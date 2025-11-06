#![cfg_attr(all(target_os = "windows", not(debug_assertions)), windows_subsystem = "windows")]

use eframe::egui::{self, RichText};
use serde_json::Value;
use std::fs;
use std::path::{Path, PathBuf};
use chardetng::EncodingDetector;

#[derive(Clone, Default)]
struct ConvertedFile {
    src: PathBuf,
    rows: Vec<Vec<Option<String>>>, // Depth_1..Depth_n + value
    max_depth: usize,
}

#[derive(Default)]
struct AppState {
    dropped: Vec<PathBuf>,
    converted: Vec<ConvertedFile>,
    last_msg: String,
    busy: bool,
    want_save: bool,
}

/* ---------- Encoding Detection ---------- */

fn detect_and_decode(bytes: &[u8]) -> anyhow::Result<String> {
    // Try UTF-8 first (most common case)
    if let Ok(text) = std::str::from_utf8(bytes) {
        // Strip BOM if present
        let text = text.strip_prefix('\u{FEFF}').unwrap_or(text);
        return Ok(text.to_string());
    }

    // Use chardetng to detect encoding
    let mut detector = EncodingDetector::new();
    detector.feed(bytes, true);
    let encoding = detector.guess(None, true);

    // Decode using the detected encoding
    let (decoded, _, had_errors) = encoding.decode(bytes);

    if had_errors {
        anyhow::bail!("Failed to decode file with detected encoding: {}", encoding.name());
    }

    // Strip BOM if present
    let decoded = decoded.strip_prefix('\u{FEFF}').unwrap_or(&decoded);
    Ok(decoded.to_string())
}

/* ---------- 파이썬 create_df 동작과 동일 ---------- */

fn count_depth(v: &Value) -> usize {
    match v {
        Value::Object(map) => {
            if map.is_empty() {
                1
            } else {
                std::cmp::max(1, map.values().map(|x| 1 + count_depth(x)).max().unwrap_or(1))
            }
        }
        Value::Array(arr) => {
            if arr.is_empty() {
                1
            } else {
                std::cmp::max(1, arr.iter().map(|x| 1 + count_depth(x)).max().unwrap_or(1))
            }
        }
        _ => 1,
    }
}

fn walk(v: &Value, path: &mut Vec<String>, rows: &mut Vec<Vec<Option<String>>>, max_depth: usize) {
    match v {
        Value::Object(map) => {
            if map.is_empty() {
                let mut row = vec![None; max_depth + 1];
                for (i, seg) in path.iter().enumerate() {
                    row[i] = Some(seg.clone());
                }
                row[max_depth] = None; // value=None
                rows.push(row);
                return;
            }
            for (k, val) in map {
                path.push(k.to_string());
                walk(val, path, rows, max_depth);
                path.pop();
            }
        }
        Value::Array(arr) => {
            if arr.is_empty() {
                let mut row = vec![None; max_depth + 1];
                for (i, seg) in path.iter().enumerate() {
                    row[i] = Some(seg.clone());
                }
                row[max_depth] = Some("[]".to_string()); // 빈 리스트 표시
                rows.push(row);
                return;
            }
            for (idx, val) in arr.iter().enumerate() {
                path.push(idx.to_string());
                walk(val, path, rows, max_depth);
                path.pop();
            }
        }
        _ => {
            let mut row = vec![None; max_depth + 1];
            for (i, seg) in path.iter().enumerate() {
                row[i] = Some(seg.clone());
            }
            let sval = match v {
                Value::String(s) => s.clone(),
                _ => v.to_string(),
            };
            row[max_depth] = Some(sval);
            rows.push(row);
        }
    }
}

fn convert_json_exact_like_py(json_text: &str) -> anyhow::Result<(Vec<Vec<Option<String>>>, usize)> {
    let v: Value = serde_json::from_str(json_text)?;
    // Calculate max depth, subtracting 1 to match Python behavior (don't count root level)
    let max_depth = count_depth(&v).saturating_sub(1);
    let mut rows = Vec::new();
    let mut path = Vec::new();
    walk(&v, &mut path, &mut rows, max_depth);
    Ok((rows, max_depth))
}

/* ---------- XLSX 저장 ---------- */

fn save_xlsx(target: &Path, rows: &[Vec<Option<String>>], max_depth: usize) -> anyhow::Result<()> {
    use rust_xlsxwriter::Workbook;

    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();

    for i in 0..max_depth {
        ws.write(0, i as u16, &format!("Depth_{}", i + 1))?;
    }
    ws.write(0, max_depth as u16, "value")?;

    for (r, row) in rows.iter().enumerate() {
        let rr = (r as u32) + 1;
        for (c, cell) in row.iter().enumerate() {
            if let Some(s) = cell {
                ws.write(rr, c as u16, s)?;
            }
        }
    }

    for i in 0..=max_depth {
        ws.set_column_width(i as u16, 16.0)?;
    }

    wb.save(target)?;
    Ok(())
}

/* ---------- GUI ---------- */

impl AppState {
    fn ui(&mut self, ctx: &egui::Context) {
        if !self.busy {
            let dropped: Vec<_> =
                ctx.input(|i| i.raw.dropped_files.iter().filter_map(|f| f.path.clone()).collect());
            if !dropped.is_empty() {
                self.dropped = dropped;
                self.converted.clear();
                self.last_msg = if self.dropped.len() == 1 {
                    format!("Dropped file: {}", self.dropped[0].display())
                } else {
                    format!(
                        "Dropped {} JSON files\nFirst: {}",
                        self.dropped.len(),
                        self.dropped[0].display()
                    )
                };
            }
        }

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.vertical_centered(|ui| {
                ui.add_space(8.0);
                ui.label(RichText::new("Drag & drop JSON files here").size(16.0));

                ui.add_space(10.0);
                ui.group(|ui| {
                    ui.set_height(120.0);
                    ui.vertical_centered(|ui| {
                        ui.label(self.last_msg.clone());
                    });
                });

                ui.add_space(8.0);
                ui.horizontal(|ui| {
                    if ui
                        .add_enabled(!self.busy, egui::Button::new("Open…"))
                        .clicked()
                    {
                        if let Some(p) =
                            rfd::FileDialog::new().add_filter("json", &["json"]).pick_files()
                        {
                            self.dropped = p;
                            self.converted.clear();
                            self.last_msg = if self.dropped.len() == 1 {
                                format!("Selected: {}", self.dropped[0].display())
                            } else {
                                format!(
                                    "Selected {} JSON files\nFirst: {}",
                                    self.dropped.len(),
                                    self.dropped[0].display()
                                )
                            };
                        }
                    }

                    let can_convert = !self.dropped.is_empty() && !self.busy;
                    if ui
                        .add_enabled(can_convert, egui::Button::new("Convert"))
                        .clicked()
                    {
                        self.converted.clear();
                        self.busy = true;
                        self.want_save = false;

                        let files = self.dropped.clone();
                        let mut converted = Vec::new();
                        let mut ok = 0usize;
                        let mut fail = 0usize;

                        for p in files {
                            match fs::read(&p) {
                                Ok(bytes) => match detect_and_decode(&bytes) {
                                    Ok(text) => match convert_json_exact_like_py(&text) {
                                        Ok((rows, depth)) => {
                                            converted.push(ConvertedFile {
                                                src: p.clone(),
                                                rows,
                                                max_depth: depth,
                                            });
                                            ok += 1;
                                        }
                                        Err(e) => {
                                            eprintln!("convert error for {}: {e}", p.display());
                                            fail += 1;
                                        }
                                    },
                                    Err(e) => {
                                        eprintln!("decode error for {}: {e}", p.display());
                                        fail += 1;
                                    }
                                },
                                Err(e) => {
                                    eprintln!("read error for {}: {e}", p.display());
                                    fail += 1;
                                }
                            }
                        }

                        self.converted = converted;
                        self.busy = false;
                        self.want_save = ok > 0;
                        self.last_msg = if ok > 0 && fail == 0 {
                            format!("Conversion complete! Files converted: {ok}")
                        } else if ok > 0 {
                            format!("Partial conversion ⚠️  Converted: {ok}, Errors: {fail}")
                        } else {
                            "Conversion failed".to_string()
                        };
                        ctx.request_repaint();
                    }

                    if ui
                        .add_enabled(self.want_save && !self.busy, egui::Button::new("Save (.xlsx)"))
                        .clicked()
                    {
                        if self.converted.len() == 1 {
                            let c = &self.converted[0];
                            let mut suggested = c.src.clone();
                            suggested.set_extension("xlsx");
                            if let Some(dest) = rfd::FileDialog::new()
                                .set_file_name(
                                    suggested
                                        .file_name()
                                        .unwrap()
                                        .to_string_lossy()
                                        .to_string(),
                                )
                                .save_file()
                            {
                                match save_xlsx(&dest, &c.rows, c.max_depth) {
                                    Ok(_) => self.last_msg = format!("Saved\n{}", dest.display()),
                                    Err(e) => self.last_msg = format!("Save failed\n{e}"),
                                }
                            } else {
                                self.last_msg = "Save cancelled.".into();
                            }
                        } else {
                            if let Some(dir) = rfd::FileDialog::new().pick_folder() {
                                let mut ok = 0usize;
                                let mut fail = 0usize;
                                for c in &self.converted {
                                    let base =
                                        format!("{}.xlsx", c.src.file_stem().unwrap().to_string_lossy());
                                    let dest = dir.join(base);
                                    match save_xlsx(&dest, &c.rows, c.max_depth) {
                                        Ok(_) => ok += 1,
                                        Err(e) => {
                                            eprintln!("save error: {e}");
                                            fail += 1;
                                        }
                                    }
                                }
                                self.last_msg = if fail == 0 {
                                    format!("Saved Files: {ok}\nFolder: {}", dir.display())
                                } else if ok == 0 {
                                    "All saves failed".into()
                                } else {
                                    format!("Partial success\nSaved: {ok}, Failed: {fail}")
                                };
                            } else {
                                self.last_msg = "Save cancelled.".into();
                            }
                        }
                        self.want_save = false;
                        ctx.request_repaint();
                    }
                });

                if self.busy {
                    ui.add_space(10.0);
                    ui.spinner();
                    ui.label("Processing…");
                }
            });
        });
    }
}

impl eframe::App for AppState {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.ui(ctx);
    }
}

/* ---------- Entry ---------- */

fn main() -> eframe::Result<()> {
    let native_options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size(egui::vec2(460.0, 280.0)),
        ..Default::default()
    };

    eframe::run_native(
        "JSON → Excel Converter",
        native_options,
        Box::new(|_cc| Ok(Box::new(AppState::default()))),
    )
}
