#![cfg_attr(
    all(target_os = "windows", not(debug_assertions)),
    windows_subsystem = "windows"
)]

mod backend;
mod gui;

use eframe::egui;
use gui::AppState;

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
