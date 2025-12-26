use crate::backend::*;
use eframe::egui::{self, RichText};
use std::fs;
use std::path::PathBuf;

#[derive(Default)]
pub struct AppState {
    dropped: Vec<PathBuf>,
    converted: Vec<ConvertedFile>,
    last_msg: String,
    busy: bool,
    want_save: bool,
}

impl AppState {
    fn ui(&mut self, ctx: &egui::Context) {
        if !self.busy {
            let dropped: Vec<_> = ctx.input(|i| {
                i.raw
                    .dropped_files
                    .iter()
                    .filter_map(|f| f.path.clone())
                    .collect()
            });
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
                        if let Some(p) = rfd::FileDialog::new()
                            .add_filter("json", &["json"])
                            .pick_files()
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
                                    Ok(text) => match create_df(&text) {
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
                        .add_enabled(
                            self.want_save && !self.busy,
                            egui::Button::new("Save (.xlsx)"),
                        )
                        .clicked()
                    {
                        if self.converted.len() == 1 {
                            let c = &self.converted[0];
                            let mut suggested = c.src.clone();
                            suggested.set_extension("xlsx");
                            if let Some(dest) = rfd::FileDialog::new()
                                .set_file_name(
                                    suggested.file_name().unwrap().to_string_lossy().to_string(),
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
                        } else if let Some(dir) = rfd::FileDialog::new().pick_folder() {
                            let mut ok = 0usize;
                            let mut fail = 0usize;
                            for c in &self.converted {
                                let base = format!(
                                    "{}.xlsx",
                                    c.src.file_stem().unwrap().to_string_lossy()
                                );
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
