use chardetng::EncodingDetector;
use serde_json::Value;
use std::path::{Path, PathBuf};

#[derive(Clone, Default)]
pub struct ConvertedFile {
    pub src: PathBuf,
    pub rows: Vec<Vec<Option<String>>>, 
    pub max_depth: usize,
}


pub fn detect_and_decode(bytes: &[u8]) -> anyhow::Result<String> {
    if let Ok(text) = std::str::from_utf8(bytes) {
        let text = text.strip_prefix('\u{FEFF}').unwrap_or(text);
        return Ok(text.to_string());
    }

    let mut detector = EncodingDetector::new();
    detector.feed(bytes, true);
    let encoding = detector.guess(None, true);

    let (decoded, _, had_errors) = encoding.decode(bytes);

    if had_errors {
        anyhow::bail!(
            "Failed to decode file with detected encoding: {}",
            encoding.name()
        );
    }

    // Strip BOM if present
    let decoded = decoded.strip_prefix('\u{FEFF}').unwrap_or(&decoded);
    Ok(decoded.to_string())
}


fn count_depth(v: &Value) -> usize {
    match v {
        Value::Object(map) => {
            if map.is_empty() {
                1
            } else {
                std::cmp::max(
                    1,
                    map.values().map(|x| 1 + count_depth(x)).max().unwrap_or(1),
                )
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
                row[max_depth] = None; 
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
                row[max_depth] = Some("[]".to_string());
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

pub fn create_df(
    json_text: &str,
) -> anyhow::Result<(Vec<Vec<Option<String>>>, usize)> {
    let v: Value = serde_json::from_str(json_text)?;
    let max_depth = count_depth(&v).saturating_sub(1);
    let mut rows = Vec::new();
    let mut path = Vec::new();
    walk(&v, &mut path, &mut rows, max_depth);
    Ok((rows, max_depth))
}

/* ---------- Save XLSX  ---------- */

pub fn save_xlsx(
    target: &Path,
    rows: &[Vec<Option<String>>],
    max_depth: usize,
) -> anyhow::Result<()> {
    use rust_xlsxwriter::Workbook;

    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();

    for i in 0..max_depth {
        ws.write(0, i as u16, format!("Depth_{}", i + 1))?;
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
