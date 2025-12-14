import sys
import math
import audioop
import tempfile
import datetime
import json
from collections import deque
from pathlib import Path
from typing import Optional

import pyaudio
from PyQt6 import QtCore, QtWidgets, QtMultimedia

import speech_recognition as sr
from gtts import gTTS
from pydub import AudioSegment

# (label, STT code, TTS code)
LANG_OPTIONS = [
    ("Ukrainian", "uk-UA", "uk"),
    ("English (US)", "en-US", "en"),
    ("Polski", "pl-PL", "pl"),
    ("Deutsch", "de-DE", "de"),
    ("Français", "fr-FR", "fr"),
    ("Español", "es-ES", "es"),
    ("Italiano", "it-IT", "it"),
    ("Português", "pt-PT", "pt"),
    ("Türkçe", "tr-TR", "tr"),
    ("Japanese", "ja-JP", "ja"),
]


def speed_change(segment: AudioSegment, factor: float) -> AudioSegment:
    # Change frame rate to alter speed without changing pitch
    new_sr = int(segment.frame_rate * factor)
    return segment._spawn(segment.raw_data, overrides={"frame_rate": new_sr}).set_frame_rate(segment.frame_rate)


def volume_gain_db(volume_factor: float) -> float:
    # Convert linear volume factor (0.1..2.0) to decibels
    volume_factor = max(0.1, min(volume_factor, 2.0))
    return 20 * math.log10(volume_factor)


class TTSWorker(QtCore.QThread):
    finished = QtCore.pyqtSignal(str)
    playback_ready = QtCore.pyqtSignal(str)
    error = QtCore.pyqtSignal(str)
    status = QtCore.pyqtSignal(str)

    def __init__(self, text: str, lang_code: str, speed_factor: float, volume_factor: float, save_path: Optional[Path]):
        super().__init__()
        self.text = text
        self.lang_code = lang_code
        self.speed_factor = speed_factor
        self.volume_factor = volume_factor
        self.save_path = save_path

    def run(self):
        try:
            self.status.emit("Синтезую...")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp_mp3:
                gtts_obj = gTTS(text=self.text, lang=self.lang_code)
                gtts_obj.save(tmp_mp3.name)

            segment = AudioSegment.from_file(tmp_mp3.name, format="mp3")
            segment = speed_change(segment, self.speed_factor)
            segment = segment + volume_gain_db(self.volume_factor)

            if self.save_path:
                segment.export(self.save_path, format="mp3")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp_wav:
                segment.export(tmp_wav.name, format="wav")
                self.playback_ready.emit(tmp_wav.name)

            self.finished.emit("Готово")
        except Exception as exc:  # noqa: BLE001
            self.error.emit(f"TTS помилка: {exc}")
        finally:
            self.status.emit("")


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Speech <-> Text Tool (PyQt6)")
        self.resize(950, 650)

        self.history = deque(maxlen=10)
        self.stt_stop = None
        self.recognizer = sr.Recognizer()
        self.config_path = self._compute_config_path()
        self.config = {}
        self.loading_config = True
        self.last_save_dir = Path.home()
        self.load_config()
        self.active_mics = []

        self.sound_effect = QtMultimedia.QSoundEffect(self)
        self.should_play = True

        self.vu_timer = QtCore.QTimer(self)
        self.vu_timer.setInterval(100)
        self.vu_timer.timeout.connect(self.update_vu_level)
        self.vu_audio = None
        self.vu_stream = None

        self.init_ui()

    def init_ui(self):
        tabs = QtWidgets.QTabWidget()

        # STT tab
        stt_tab = QtWidgets.QWidget()
        stt_layout = QtWidgets.QVBoxLayout(stt_tab)

        lang_box = QtWidgets.QHBoxLayout()
        lang_box.addWidget(QtWidgets.QLabel("Мова розпізнавання:"))
        self.stt_lang = QtWidgets.QComboBox()
        for title, stt_code, _ in LANG_OPTIONS:
            self.stt_lang.addItem(title, stt_code)
        lang_box.addWidget(self.stt_lang)
        lang_box.addStretch()
        stt_layout.addLayout(lang_box)

        mic_box = QtWidgets.QHBoxLayout()
        mic_box.addWidget(QtWidgets.QLabel("Мікрофон:"))
        self.mic_combo = QtWidgets.QComboBox()
        mic_box.addWidget(self.mic_combo, 1)
        refresh_mics = QtWidgets.QPushButton("Оновити")
        refresh_mics.clicked.connect(self.populate_microphones)
        mic_box.addWidget(refresh_mics)
        mic_box.addStretch()
        stt_layout.addLayout(mic_box)

        self.stt_text = QtWidgets.QPlainTextEdit()
        self.stt_text.setPlaceholderText("Сказане текстом...")
        stt_layout.addWidget(self.stt_text)

        btn_box = QtWidgets.QHBoxLayout()
        self.btn_start = QtWidgets.QPushButton("Почати запис")
        self.btn_stop = QtWidgets.QPushButton("Зупинити запис")
        self.btn_stop.setEnabled(False)
        btn_box.addWidget(self.btn_start)
        btn_box.addWidget(self.btn_stop)
        self.btn_copy_stt = QtWidgets.QPushButton("Скопіювати текст")
        btn_box.addWidget(self.btn_copy_stt)
        btn_box.addStretch()
        stt_layout.addLayout(btn_box)

        self.stt_status = QtWidgets.QLabel("")
        self.stt_progress = QtWidgets.QProgressBar()
        self.stt_progress.setRange(0, 0)  # Indeterminate indicator
        self.stt_progress.setVisible(False)

        vu_box = QtWidgets.QHBoxLayout()
        vu_box.addWidget(QtWidgets.QLabel("Рівень сигналу:"))
        self.vu_meter = QtWidgets.QProgressBar()
        self.vu_meter.setRange(0, 100)
        self.vu_meter.setTextVisible(False)
        vu_box.addWidget(self.vu_meter, 1)
        vu_box.addStretch()

        stt_layout.addWidget(self.stt_status)
        stt_layout.addWidget(self.stt_progress)
        stt_layout.addLayout(vu_box)

        tabs.addTab(stt_tab, "Speech -> Text")

        # TTS tab
        tts_tab = QtWidgets.QWidget()
        tts_layout = QtWidgets.QVBoxLayout(tts_tab)

        tts_lang_box = QtWidgets.QHBoxLayout()
        tts_lang_box.addWidget(QtWidgets.QLabel("Мова синтезу:"))
        self.tts_lang = QtWidgets.QComboBox()
        for title, _, tts_code in LANG_OPTIONS:
            self.tts_lang.addItem(title, tts_code)
        tts_lang_box.addWidget(self.tts_lang)
        tts_lang_box.addStretch()
        tts_layout.addLayout(tts_lang_box)

        self.tts_text = QtWidgets.QPlainTextEdit()
        self.tts_text.setPlaceholderText("Введіть текст для озвучення...")
        tts_layout.addWidget(self.tts_text)

        controls = QtWidgets.QHBoxLayout()
        controls.addWidget(QtWidgets.QLabel("Швидкість:"))
        self.speed_combo = QtWidgets.QComboBox()
        for factor in ["0.5", "1", "1.5", "2", "2.5"]:
            self.speed_combo.addItem(f"x{factor}", float(factor))
        self.speed_combo.setCurrentIndex(1)
        controls.addWidget(self.speed_combo)

        controls.addWidget(QtWidgets.QLabel("Гучність (%):"))
        self.volume_slider = QtWidgets.QSlider(QtCore.Qt.Orientation.Horizontal)
        self.volume_slider.setMinimum(20)
        self.volume_slider.setMaximum(200)
        self.volume_slider.setValue(100)
        controls.addWidget(self.volume_slider)
        controls.addStretch()
        tts_layout.addLayout(controls)

        tts_btns = QtWidgets.QHBoxLayout()
        self.btn_play = QtWidgets.QPushButton("Відтворити")
        self.btn_save_mp3 = QtWidgets.QPushButton("Зберегти MP3")
        tts_btns.addWidget(self.btn_play)
        tts_btns.addWidget(self.btn_save_mp3)
        tts_btns.addStretch()
        tts_layout.addLayout(tts_btns)

        self.tts_status = QtWidgets.QLabel("")
        self.tts_progress = QtWidgets.QProgressBar()
        self.tts_progress.setRange(0, 0)
        self.tts_progress.setVisible(False)
        tts_layout.addWidget(self.tts_status)
        tts_layout.addWidget(self.tts_progress)

        tabs.addTab(tts_tab, "Text -> Speech")

        # History + export
        bottom = QtWidgets.QHBoxLayout()
        self.history_list = QtWidgets.QListWidget()
        bottom.addWidget(self.history_list, 2)

        export_box = QtWidgets.QVBoxLayout()
        self.btn_export = QtWidgets.QPushButton("Експортувати в TXT")
        export_box.addWidget(self.btn_export)
        export_box.addStretch()
        bottom.addLayout(export_box, 1)

        root_layout = QtWidgets.QVBoxLayout()
        root_layout.addWidget(tabs)
        root_layout.addLayout(bottom)

        wrapper = QtWidgets.QWidget()
        wrapper.setLayout(root_layout)
        self.setCentralWidget(wrapper)

        # Signals
        self.btn_start.clicked.connect(self.start_recording)
        self.btn_stop.clicked.connect(self.stop_recording)
        self.btn_play.clicked.connect(self.play_tts)
        self.btn_save_mp3.clicked.connect(lambda: self.play_tts(save_only=True))
        self.btn_export.clicked.connect(self.export_texts)
        self.btn_copy_stt.clicked.connect(self.copy_stt_text)
        self.stt_lang.currentIndexChanged.connect(self.save_settings)
        self.tts_lang.currentIndexChanged.connect(self.save_settings)
        self.speed_combo.currentIndexChanged.connect(self.save_settings)
        self.volume_slider.valueChanged.connect(self.save_settings)
        self.mic_combo.currentIndexChanged.connect(self.save_settings)

        self.populate_microphones()
        self.restore_settings()

    # ---------- STT ----------
    def start_recording(self):
        if self.stt_stop:
            return
        lang_code = self.stt_lang.currentData()
        mic_index = self.mic_combo.currentData()
        if mic_index is None or not self.active_mics:
            QtWidgets.QMessageBox.critical(self, "Помилка", "Не знайдено активного мікрофона.")
            return
        try:
            mic_index_int = int(mic_index)
            mic = sr.Microphone(device_index=mic_index_int)
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Помилка", f"Помилка ініціалізації мікрофона: {exc}")
            return

        self.stt_status.setText("Слухаю...")
        self.stt_progress.setVisible(True)
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.start_vu_meter(mic_index_int)

        def callback(_, audio):
            try:
                text = self.recognizer.recognize_google(audio, language=lang_code)
                self.stt_text.appendPlainText(text)
                self.add_history(f"STT [{lang_code}]: {text}")
            except sr.UnknownValueError:
                self.add_history("STT: не вдалося розпізнати фразу")
            except sr.RequestError:
                self.add_history("STT: проблема з підключенням до сервісу")
            except Exception as exc_inner:  # noqa: BLE001
                self.add_history(f"STT помилка: {exc_inner}")

        self.stt_stop = self.recognizer.listen_in_background(mic, callback)

    def stop_recording(self):
        if self.stt_stop:
            try:
                self.stt_stop(wait_for_stop=False)
            except Exception:
                pass
        self.stt_stop = None
        self.stt_status.setText("Зупинено")
        self.stt_progress.setVisible(False)
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.stop_vu_meter()

    # ---------- TTS ----------
    def play_tts(self, save_only: bool = False):
        text = self.tts_text.toPlainText().strip()
        if not text:
            QtWidgets.QMessageBox.warning(self, "Увага", "Введіть текст для озвучення.")
            return

        lang_code = self.tts_lang.currentData()
        speed_factor = float(self.speed_combo.currentData())
        volume_factor = self.volume_slider.value() / 100.0

        save_path = None
        if save_only:
            target, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Зберегти MP3", str(self.last_save_dir / "output.mp3"), "MP3 Files (*.mp3)")
            if not target:
                return
            save_path = Path(target)
            self.last_save_dir = Path(target).parent
            self.save_settings()

        self.tts_status.setText("Готую аудіо...")
        self.tts_progress.setVisible(True)
        self.btn_play.setEnabled(False)
        self.btn_save_mp3.setEnabled(False)

        self.should_play = not save_only
        self.tts_worker = TTSWorker(text, lang_code, speed_factor, volume_factor, save_path)
        self.tts_worker.status.connect(self.tts_status.setText)
        self.tts_worker.finished.connect(self.on_tts_done)
        self.tts_worker.error.connect(self.on_tts_error)
        self.tts_worker.playback_ready.connect(self.on_playback_ready)
        self.tts_worker.start()

    def on_playback_ready(self, wav_path: str):
        if not self.should_play:
            return
        self.sound_effect.stop()
        self.sound_effect.setSource(QtCore.QUrl.fromLocalFile(wav_path))
        self.sound_effect.setLoopCount(1)
        vol = max(0, min(100, self.volume_slider.value()))
        self.sound_effect.setVolume(vol / 100.0)
        self.sound_effect.play()

    def on_tts_done(self, msg: str):
        self.tts_status.setText(msg)
        self.tts_progress.setVisible(False)
        self.btn_play.setEnabled(True)
        self.btn_save_mp3.setEnabled(True)
        self.add_history(f"TTS [{self.tts_lang.currentData()}]: {self.tts_text.toPlainText().strip()}")

    def on_tts_error(self, msg: str):
        QtWidgets.QMessageBox.critical(self, "TTS помилка", msg)
        self.tts_status.setText("Помилка")
        self.tts_progress.setVisible(False)
        self.btn_play.setEnabled(True)
        self.btn_save_mp3.setEnabled(True)

    # ---------- History & Export ----------
    def add_history(self, entry: str):
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {entry}"
        self.history.appendleft(line)
        self.history_list.clear()
        self.history_list.addItems(self.history)

    def export_texts(self):
        target, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Зберегти результати", str(self.last_save_dir / "result.txt"), "Text Files (*.txt)")
        if not target:
            return
        try:
            with open(target, "w", encoding="utf-8") as f:
                f.write("=== Speech-to-Text ===\n")
                f.write(self.stt_text.toPlainText().strip() + "\n\n")
                f.write("=== Text-to-Speech ===\n")
                f.write(self.tts_text.toPlainText().strip() + "\n\n")
                f.write("=== Історія ===\n")
                for item in self.history:
                    f.write(item + "\n")
            QtWidgets.QMessageBox.information(self, "Успішно", "Файл успішно збережено.")
            self.last_save_dir = Path(target).parent
            self.save_settings()
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Помилка", f"Не вдалося зберегти файл: {exc}")

    # ---------- Mic & VU ----------
    def populate_microphones(self):
        self.mic_combo.clear()
        self.active_mics = []
        pa = None
        try:
            pa = pyaudio.PyAudio()
            for idx in range(pa.get_device_count()):
                info = pa.get_device_info_by_index(idx)
                if info.get("maxInputChannels", 0) <= 0:
                    continue
                try:
                    pa.is_format_supported(
                        rate=int(info.get("defaultSampleRate", 16000)),
                        input_device=idx,
                        input_channels=1,
                        input_format=pyaudio.paInt16,
                    )
                except Exception:
                    continue
                name = info.get("name", f"Device {idx}")
                self.active_mics.append((idx, name))
        except Exception:
            self.active_mics = []
        finally:
            if pa:
                try:
                    pa.terminate()
                except Exception:
                    pass

        for idx, name in self.active_mics:
            self.mic_combo.addItem(f"{idx}: {name}", idx)
        if not self.active_mics:
            self.mic_combo.addItem("Немає пристроїв", None)
            return
        saved_index = self.config.get("audio_mic_index")
        try:
            if saved_index is not None:
                index_int = int(saved_index)
            else:
                index_int = None
        except (TypeError, ValueError):
            index_int = None
        if index_int is not None:
            combo_index = self.mic_combo.findData(index_int)
            if combo_index >= 0:
                self.mic_combo.setCurrentIndex(combo_index)
        if self.mic_combo.currentIndex() < 0 and self.mic_combo.count() > 0:
            self.mic_combo.setCurrentIndex(0)

    def start_vu_meter(self, device_index: Optional[int]):
        self.stop_vu_meter()
        if device_index is None:
            return
        try:
            self.vu_audio = pyaudio.PyAudio()
            device_info = self.vu_audio.get_device_info_by_index(device_index)
            rate = int(device_info.get("defaultSampleRate", 16000))
            self.vu_stream = self.vu_audio.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=rate,
                input=True,
                input_device_index=device_index,
                frames_per_buffer=1024,
            )
            self.vu_timer.start()
        except Exception:
            self.vu_meter.setValue(0)

    def stop_vu_meter(self):
        self.vu_timer.stop()
        if self.vu_stream:
            try:
                self.vu_stream.stop_stream()
                self.vu_stream.close()
            except Exception:
                pass
        self.vu_stream = None
        if self.vu_audio:
            try:
                self.vu_audio.terminate()
            except Exception:
                pass
        self.vu_audio = None
        self.vu_meter.setValue(0)

    def update_vu_level(self):
        if not self.vu_stream:
            self.vu_meter.setValue(0)
            return
        try:
            data = self.vu_stream.read(1024, exception_on_overflow=False)
            rms = audioop.rms(data, 2)
            level = min(100, int(rms / 300))
            self.vu_meter.setValue(level)
        except Exception:
            self.vu_meter.setValue(0)

    # ---------- Clipboard ----------
    def copy_stt_text(self):
        text = self.stt_text.toPlainText().strip()
        if not text:
            QtWidgets.QMessageBox.information(self, "Копіювання", "Немає тексту для копіювання.")
            return
        QtWidgets.QApplication.clipboard().setText(text)
        self.add_history("STT текст скопійовано у буфер обміну.")

    # ---------- Settings ----------
    def _compute_config_path(self) -> Path:
        if getattr(sys, "frozen", False):
            base_dir = Path(sys.executable).resolve().parent
        else:
            base_dir = Path(__file__).resolve().parent
        return base_dir / "config.json"

    # JSON config helpers
    def load_config(self):
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                self.config = data
        except FileNotFoundError:
            self.config = {}
        except Exception:
            self.config = {}
        self.last_save_dir = Path(self.config.get("paths_last_save_dir", str(Path.home())))

    def _read_int_config(self, key: str, default: int = 0) -> int:
        try:
            return int(self.config.get(key, default))
        except (TypeError, ValueError):
            return default

    def _write_config_file(self):
        try:
            self.config_path.write_text(json.dumps(self.config, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def restore_settings(self):
        stt_idx = self._read_int_config("lang_stt_index", 0)
        tts_idx = self._read_int_config("lang_tts_index", 0)
        speed_idx = self._read_int_config("tts_speed_index", 1)
        volume = self._read_int_config("tts_volume", 100)
        self.last_save_dir = Path(self.config.get("paths_last_save_dir", str(Path.home())))
        if 0 <= stt_idx < self.stt_lang.count():
            self.stt_lang.setCurrentIndex(stt_idx)
        if 0 <= tts_idx < self.tts_lang.count():
            self.tts_lang.setCurrentIndex(tts_idx)
        if 0 <= speed_idx < self.speed_combo.count():
            self.speed_combo.setCurrentIndex(speed_idx)
        self.volume_slider.setValue(volume)
        saved_mic = self.config.get("audio_mic_index")
        try:
            if saved_mic is not None:
                saved_idx = int(saved_mic)
                combo_index = self.mic_combo.findData(saved_idx)
                if combo_index >= 0:
                    self.mic_combo.setCurrentIndex(combo_index)
        except (TypeError, ValueError):
            pass
        # finished initial load; allow saves after this
        self.loading_config = False

    def save_settings(self):
        if getattr(self, "loading_config", False):
            return
        self.config["lang_stt_index"] = self.stt_lang.currentIndex()
        self.config["lang_tts_index"] = self.tts_lang.currentIndex()
        self.config["tts_speed_index"] = self.speed_combo.currentIndex()
        self.config["tts_volume"] = self.volume_slider.value()
        self.config["paths_last_save_dir"] = str(self.last_save_dir)
        mic_data = self.mic_combo.currentData()
        if mic_data is not None:
            try:
                self.config["audio_mic_index"] = int(mic_data)
            except (TypeError, ValueError):
                pass
        self._write_config_file()

    def closeEvent(self, event):  # noqa: N802
        self.save_settings()
        super().closeEvent(event)


def main():
    app = QtWidgets.QApplication(sys.argv)
    QtCore.QCoreApplication.setOrganizationName("CourseWork")
    QtCore.QCoreApplication.setApplicationName("SpeechTextTool")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
