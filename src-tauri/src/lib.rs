//! Study Pacer for Windows: the same app as the website in its own small window, reading your documents
//! straight from folders on the computer. Nothing is copied: the page asks for a folder's file list (scan) and
//! reads each file through Tauri's asset protocol when it needs it.

use serde::Serialize;
use std::fs;
use std::path::Path;
use std::time::UNIX_EPOCH;
use tauri::AppHandle;
use tauri_plugin_dialog::DialogExt;

const EXTS: &[&str] = &[
    "pdf", "pptx", "docx", "html", "htm", "xhtml", "epub", "md", "markdown", "txt", "text", "png", "jpg", "jpeg", "webp", "gif", "avif",
];
const MAX_FILES: usize = 25_000;
const MAX_DEPTH: usize = 16;

#[derive(Serialize)]
struct Entry {
    path: String,
    rel: String,
    size: u64,
    mtime: f64,
}

fn walk(dir: &Path, root: &Path, depth: usize, out: &mut Vec<Entry>) {
    if depth > MAX_DEPTH || out.len() >= MAX_FILES {
        return;
    }
    let Ok(rd) = fs::read_dir(dir) else { return };
    for e in rd.flatten() {
        let name = e.file_name().to_string_lossy().to_string();
        // hidden things, Office's lock files and system folders aren't study material
        if name.starts_with('.') || name.starts_with("~$") || name.eq_ignore_ascii_case("node_modules") || name.eq_ignore_ascii_case("$RECYCLE.BIN") {
            continue;
        }
        let Ok(ft) = e.file_type() else { continue };
        let p = e.path();
        if ft.is_dir() {
            walk(&p, root, depth + 1, out);
        } else if ft.is_file() {
            let ext = p.extension().map(|x| x.to_string_lossy().to_lowercase()).unwrap_or_default();
            if !EXTS.contains(&ext.as_str()) {
                continue;
            }
            let md = e.metadata().ok();
            let size = md.as_ref().map(|m| m.len()).unwrap_or(0);
            let mtime = md
                .and_then(|m| m.modified().ok())
                .and_then(|t| t.duration_since(UNIX_EPOCH).ok())
                .map(|d| d.as_millis() as f64)
                .unwrap_or(0.0);
            let rel = p.strip_prefix(root).unwrap_or(&p).to_string_lossy().replace('\\', "/");
            out.push(Entry { path: p.to_string_lossy().to_string(), rel, size, mtime });
            if out.len() >= MAX_FILES {
                return;
            }
        }
    }
}

/// Every readable document under a folder, at any depth, with its path relative to that folder.
#[tauri::command]
async fn scan(root: String) -> Result<Vec<Entry>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let r = Path::new(&root);
        if !r.is_dir() {
            return Err("missing".to_string());
        }
        let mut out = Vec::new();
        walk(r, r, 0, &mut out);
        Ok(out)
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Choose a folder to add to the library.
#[tauri::command]
async fn pick_folder(app: AppHandle) -> Result<Option<String>, String> {
    let picked = app.dialog().file().set_title("Add a folder of documents to Study Pacer").blocking_pick_folder();
    let Some(fp) = picked else { return Ok(None) };
    let p = fp.into_path().map_err(|e| e.to_string())?;
    Ok(Some(p.to_string_lossy().to_string()))
}

/// Show a file (or folder) in File Explorer.
#[tauri::command]
fn reveal(path: String) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        let p = Path::new(&path);
        let mut cmd = std::process::Command::new("explorer");
        if p.is_file() { cmd.arg(format!("/select,{}", path)); } else { cmd.arg(&path); }
        cmd.spawn().map_err(|e| e.to_string())?;
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = path;
    }
    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![scan, pick_folder, reveal])
        .run(tauri::generate_context!())
        .expect("error while running Study Pacer");
}
