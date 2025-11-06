import sys, os, json
from typing import TYPE_CHECKING, List, Tuple
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication, QLabel, QVBoxLayout, QWidget, QPushButton,
    QFileDialog, QFrame, QHBoxLayout
)

if TYPE_CHECKING:
    import pandas as pd


class LoadingOverlay(QWidget):
    def __init__(self, parent: QWidget, text: str = "Loading..."):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        frame = QFrame(self)
        frame.setStyleSheet("""
            QFrame { background: rgba(30,30,30,200); border-radius: 12px; }
            QLabel { color: #EEE; font-size: 13px; }
        """)
        layout = QVBoxLayout(frame)
        t = QLabel("Initializing…"); t.setAlignment(Qt.AlignmentFlag.AlignCenter)
        m = QLabel(text); m.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(t); layout.addWidget(m)
        box = QHBoxLayout(self)
        box.addWidget(frame, alignment=Qt.AlignmentFlag.AlignCenter)
        self.resize_to_parent()

    def resize_to_parent(self):
        if self.parent():
            self.setGeometry(self.parent().rect())

    def showEvent(self, e): super().showEvent(e); self.resize_to_parent()
    def resizeEvent(self, e): super().resizeEvent(e); self.resize_to_parent()


class WarmupThread(QThread):
    finished_ok = pyqtSignal()
    failed = pyqtSignal(str)

    def run(self):
        try:
            import pandas, openpyxl
            self.finished_ok.emit()
        except Exception as e:
            self.failed.emit(str(e))


class MainWindows(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Drag & Drop")
        self.setGeometry(100, 100, 420, 220)
        self.dropped_paths: List[str] = []
        self.last_path: str | None = None
        self.converted: List[Tuple[str, "pd.DataFrame"]] = []

        self.label = QLabel("Drag & drop JSON files here")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setFont(QFont("", 11))
        self.label.setStyleSheet("QLabel { border:2px dashed #888; border-radius:10px; padding:20px; color:#555; }")

        self.convert_btn = QPushButton("Convert")
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.on_main_clicked)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.convert_btn)
        self.setLayout(layout)
        self.setAcceptDrops(True)

        self.overlay = LoadingOverlay(self, "Loading pandas/openpyxl…")
        self.overlay.show()

        self.warmup = WarmupThread()
        self.warmup.finished_ok.connect(lambda: (self.overlay.hide(), self.convert_btn.setEnabled(True)))
        self.warmup.failed.connect(lambda msg: (self.overlay.hide(), self.convert_btn.setEnabled(True), self.label.setText(f"Warmup failed:\n{msg}")))
        self.warmup.start()

    def dragEnterEvent(self, e):
        m = e.mimeData()
        if m.hasUrls() and any(u.isLocalFile() and u.toLocalFile().lower().endswith(".json") for u in m.urls()):
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e):
        urls = [u for u in e.mimeData().urls() if u.isLocalFile()]
        self.dropped_paths = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(".json")]
        if not self.dropped_paths:
            self.label.setText("Only JSON files are supported."); self.convert_btn.setEnabled(False); e.ignore(); return
        self.last_path = self.dropped_paths[0]
        self.converted = []
        self.convert_btn.setText("Convert")
        self.convert_btn.setEnabled(True)
        msg = f"Dropped {len(self.dropped_paths)} JSON files.\nFirst: {self.last_path}" if len(self.dropped_paths) > 1 else f"Dropped file:\n{self.last_path}"
        self.label.setText(msg)
        e.acceptProposedAction()

    def on_main_clicked(self):
        if self.convert_btn.text() == "Convert": self.run_convert_only()
        else: self.run_download_only()

    def run_convert_only(self):
        if not self.dropped_paths: self.label.setText("Only JSON files are supported."); return
        import pandas as pd
        self.converted = []
        ok, fail = 0, 0
        for src in self.dropped_paths:
            try:
                with open(src, "r", encoding="utf-8") as f: data = json.load(f)
                df = self.create_df(data)
                self.converted.append((src, df)); ok += 1
            except Exception as e:
                fail += 1; print("Convert error:", e)
        if ok and not fail:
            self.label.setText(f"Conversion complete!\nFiles converted: {ok}")
            self.convert_btn.setText("Download")
        elif ok:
            self.label.setText(f"Partial conversion ⚠️  Converted: {ok}, Errors: {fail}")
            self.convert_btn.setText("Download")
        else:
            self.label.setText("Conversion failed")

    def run_download_only(self):
        if not self.converted: self.label.setText("Nothing to download. Convert first."); return
        import pandas as pd
        if len(self.converted) == 1:
            src, df = self.converted[0]
            suggested = os.path.splitext(src)[0] + ".xlsx"
            dest, _ = QFileDialog.getSaveFileName(self, "Save as Excel", suggested, "Excel Files (*.xlsx);;All Files (*)")
            if not dest: self.label.setText("Save cancelled."); return
            try:
                df.to_excel(dest, index=False)
                self.label.setText(f"Saved\n{dest}")
            except Exception as e:
                self.label.setText(f"Save failed\n{e}")
        else:
            target_dir = QFileDialog.getExistingDirectory(self, "Select target folder")
            if not target_dir: self.label.setText("Save cancelled."); return
            ok, fail = 0, 0
            for src, df in self.converted:
                try:
                    base = os.path.splitext(os.path.basename(src))[0] + ".xlsx"
                    df.to_excel(os.path.join(target_dir, base), index=False); ok += 1
                except Exception as e:
                    fail += 1; print("Save error:", src, e)
            self.label.setText(f"Saved Files: {ok}\nFolder: {target_dir}" if not fail else f"Partial success\nSaved: {ok}, Failed: {fail}")
        self.convert_btn.setText("Convert")

    def count_depth(self, o):
        if isinstance(o, dict): return 1 if not o else max(1, max(1 + self.count_depth(v) for v in o.values()))
        if isinstance(o, list): return 1 if not o else max(1, max(1 + self.count_depth(v) for v in o))
        return 1

    def _walk(self, o, p, rows, d):
        if isinstance(o, dict):
            if not o:
                rows.append({f"Depth_{i+1}": (p[i] if i < len(p) else "") for i in range(d)} | {"value": None}); return
            for k, v in o.items(): self._walk(v, p + [str(k)], rows, d)
        elif isinstance(o, list):
            if not o:
                rows.append({f"Depth_{i+1}": (p[i] if i < len(p) else "") for i in range(d)} | {"value": []}); return
            for i, v in enumerate(o): self._walk(v, p + [str(i)], rows, d)
        else:
            rows.append({f"Depth_{i+1}": (p[i] if i < len(p) else "") for i in range(d)} | {"value": o})

    def create_df(self, data):
        import pandas as pd
        import openpyxl
        d = self.count_depth(data)
        rows = []; self._walk(data, [], rows, d)
        return pd.DataFrame(rows, columns=[f"Depth_{i}" for i in range(1, d + 1)] + ["value"])


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindows()
    w.show()
    sys.exit(app.exec())
