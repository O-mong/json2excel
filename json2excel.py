import sys, os, json
import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget, QPushButton, QFileDialog

class MainWindows(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Drag & Drop")
        self.setGeometry(100, 100, 420, 220)

        # state
        self.dropped_paths: list[str] = []
        self.last_path: str | None = None
        self.converted: list[tuple[str, pd.DataFrame]] = []  # (src_path, df)

        # UI
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

    # ---------- Drag & Drop ----------
    def dragEnterEvent(self, event):
        mime = event.mimeData()
        if mime.hasUrls() and any(u.isLocalFile() and u.toLocalFile().lower().endswith(".json") for u in mime.urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = [u for u in event.mimeData().urls() if u.isLocalFile()]
        self.dropped_paths = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(".json")]
        if not self.dropped_paths:
            self.label.setText("Only JSON files are supported.")
            self.convert_btn.setEnabled(False)
            event.ignore()
            return

        self.last_path = self.dropped_paths[0]
        self.converted = []  # reset previous results
        self.convert_btn.setText("Convert")
        self.convert_btn.setEnabled(True)

        if len(self.dropped_paths) == 1:
            self.label.setText(f"Dropped file:\n{self.last_path}")
        else:
            self.label.setText(f"Dropped {len(self.dropped_paths)} JSON files.\nFirst: {self.last_path}")
        event.acceptProposedAction()

    # ---------- Button handler (toggle) ----------
    def on_main_clicked(self):
        if self.convert_btn.text() == "Convert":
            self.run_convert_only()
        else:  # "Download"
            self.run_download_only()

    # ---------- Convert (no saving) ----------
    def run_convert_only(self):
        if not self.dropped_paths:
            self.label.setText("Only JSON files are supported.")
            return
        self.converted = []
        ok, fail = 0, 0
        try:
            for src in self.dropped_paths:
                with open(src, "r", encoding="utf-8") as f:
                    data = json.load(f)
                df = self.create_df(data)
                self.converted.append((src, df))
                ok += 1
        except Exception as e:
            fail += 1
            print("Convert error:", e)

        if ok and not fail:
            self.label.setText(f"Conversion complete!\nFiles converted: {ok}")
            self.convert_btn.setText("Download")
        elif ok and fail:
            self.label.setText(f"Partial conversion ⚠️  Converted: {ok}, Errors: {fail}")
            self.convert_btn.setText("Download")  # still allow saving converted ones
        else:
            self.label.setText("Conversion failed")

    # ---------- Download (uses dialogs) ----------
    def run_download_only(self):
        if not self.converted:
            self.label.setText("Nothing to download. Convert first.")
            return

        # single vs multiple
        if len(self.converted) == 1:
            src, df = self.converted[0]
            suggested = os.path.splitext(src)[0] + ".xlsx"
            dest, _ = QFileDialog.getSaveFileName(self, "Save as Excel", suggested,
                                                  "Excel Files (*.xlsx);;All Files (*)")
            if not dest:
                self.label.setText("Save cancelled.")
                return
            try:
                df.to_excel(dest, index=False)
                self.label.setText(f"Saved\n{dest}")
            except Exception as e:
                self.label.setText(f"Save failed\n{e}")
        else:
            target_dir = QFileDialog.getExistingDirectory(self, "Select target folder")
            if not target_dir:
                self.label.setText("Save cancelled.")
                return
            ok, fail = 0, 0
            for src, df in self.converted:
                try:
                    base = os.path.splitext(os.path.basename(src))[0] + ".xlsx"
                    df.to_excel(os.path.join(target_dir, base), index=False)
                    ok += 1
                except Exception as e:
                    print("Save error:", src, e)
                    fail += 1
            if fail == 0:
                self.label.setText(f"Saved Files: {ok}\nFolder: {target_dir}")
            elif ok == 0:
                self.label.setText("All saves failed")
            else:
                self.label.setText(f"Partial success\nSaved: {ok}, Failed: {fail}")

        # after download, reset to Convert
        self.convert_btn.setText("Convert")

    # ---------- JSON -> DataFrame ----------
    def count_depth(self, obj):
        if isinstance(obj, dict):
            return 1 if not obj else max(1, max(1 + self.count_depth(v) for v in obj.values()))
        if isinstance(obj, list):
            return 1 if not obj else max(1, max(1 + self.count_depth(v) for v in obj))
        return 1

    def _walk(self, obj, path, rows, max_depth):
        if isinstance(obj, dict):
            if not obj:
                row = {f"Depth_{i+1}": (path[i] if i < len(path) else "") for i in range(max_depth)}
                row["value"] = None
                rows.append(row); return
            for k, v in obj.items():
                self._walk(v, path + [str(k)], rows, max_depth)
        elif isinstance(obj, list):
            if not obj:
                row = {f"Depth_{i+1}": (path[i] if i < len(path) else "") for i in range(max_depth)}
                row["value"] = []; rows.append(row); return
            for idx, v in enumerate(obj):
                self._walk(v, path + [str(idx)], rows, max_depth)
        else:
            row = {f"Depth_{i+1}": (path[i] if i < len(path) else "") for i in range(max_depth)}
            row["value"] = obj; rows.append(row)

    def create_df(self, data):
        max_depth = self.count_depth(data)
        cols = [f"Depth_{i}" for i in range(1, max_depth + 1)] + ["value"]
        rows = []; self._walk(data, [], rows, max_depth)
        return pd.DataFrame(rows, columns=cols)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindows()
    window.show()
    sys.exit(app.exec())
