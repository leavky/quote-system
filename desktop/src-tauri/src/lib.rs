use serde::Serialize;
use std::path::{Path, PathBuf};
use tauri::Manager;

#[derive(Debug, Serialize)]
struct DefaultPaths {
    ledger_path: String,
    price_path: String,
    quote_price_path: String,
    quote_template_path: String,
}

fn external_data_dirs() -> Vec<PathBuf> {
    let mut dirs = Vec::new();
    if let Ok(executable) = std::env::current_exe() {
        if let Some(executable_dir) = executable.parent() {
            dirs.push(executable_dir.to_path_buf());

            #[cfg(target_os = "macos")]
            if let Some(app_parent) = executable_dir
                .parent()
                .and_then(Path::parent)
                .and_then(Path::parent)
            {
                dirs.push(app_parent.to_path_buf());
            }
        }
    }
    if let Ok(current_dir) = std::env::current_dir() {
        if !dirs.contains(&current_dir) {
            dirs.push(current_dir);
        }
    }
    dirs
}

fn resolve_default_file(
    filename: &str,
    external_dirs: &[PathBuf],
    resource_dir: Option<&Path>,
    development_path: PathBuf,
) -> PathBuf {
    let resolved = external_dirs
        .iter()
        .map(|dir| dir.join(filename))
        .find(|path| path.is_file())
        .or_else(|| {
            resource_dir
                .map(|dir| dir.join(filename))
                .filter(|path| path.is_file())
        })
        .unwrap_or(development_path);
    if resolved.is_file() {
        resolved
    } else {
        PathBuf::new()
    }
}

#[tauri::command]
fn default_paths(app: tauri::AppHandle) -> Result<DefaultPaths, String> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .ok_or_else(|| "无法定位项目目录".to_string())?;
    let resource_dir = app.path().resource_dir().ok();
    let external_dirs = external_data_dirs();
    let ledger_path = resolve_default_file(
        "zc检测费用台帐明细20260725.xls",
        &external_dirs,
        resource_dir.as_deref(),
        root.join("data").join("zc检测费用台帐明细20260725.xls"),
    );
    let price_path = resolve_default_file(
        "检测项目结算计费价格汇总.xlsx",
        &external_dirs,
        resource_dir.as_deref(),
        root.join("data").join("检测项目结算计费价格汇总.xlsx"),
    );
    let quote_price_path = resolve_default_file(
        "检测报价表.xlsx",
        &external_dirs,
        resource_dir.as_deref(),
        root.join("data").join("检测报价表.xlsx"),
    );
    let quote_template_path = resolve_default_file(
        "报价清单导入模板.xlsx",
        &external_dirs,
        resource_dir.as_deref(),
        root.join("data").join("报价清单导入模板.xlsx"),
    );
    Ok(DefaultPaths {
        ledger_path: ledger_path.to_string_lossy().to_string(),
        price_path: price_path.to_string_lossy().to_string(),
        quote_price_path: quote_price_path.to_string_lossy().to_string(),
        quote_template_path: quote_template_path.to_string_lossy().to_string(),
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![default_paths])
        .setup(|app| {
            if cfg!(debug_assertions) {
                app.handle().plugin(
                    tauri_plugin_log::Builder::default()
                        .level(log::LevelFilter::Info)
                        .build(),
                )?;
            }
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
