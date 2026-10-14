//! Native file/folder pickers (the Windows common item dialogs).

use std::path::PathBuf;

use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance, CoTaskMemFree};
use windows::Win32::UI::Shell::{
    FOS_ALLOWMULTISELECT, FOS_FORCEFILESYSTEM, FOS_PICKFOLDERS, FileOpenDialog, FileSaveDialog,
    IFileDialog, IFileOpenDialog, IFileSaveDialog, IShellItem, SIGDN_FILESYSPATH,
};
use windows::core::{PCWSTR, Result};

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(Some(0)).collect()
}

fn path_of(item: &IShellItem) -> Result<PathBuf> {
    // SAFETY: GetDisplayName returns a CoTaskMem string we free after copying it.
    unsafe {
        let name = item.GetDisplayName(SIGDN_FILESYSPATH)?;
        let s = name.to_string().unwrap_or_default();
        CoTaskMemFree(Some(name.0 as *const _));
        Ok(PathBuf::from(s))
    }
}

fn open(owner: HWND, title: &str, folders: bool, multi: bool) -> Vec<PathBuf> {
    let run = || -> Result<Vec<PathBuf>> {
        // SAFETY: standard COM usage on the Excel UI thread (an STA).
        unsafe {
            let dlg: IFileOpenDialog =
                CoCreateInstance(&FileOpenDialog, None, CLSCTX_INPROC_SERVER)?;
            let mut opts = dlg.GetOptions()? | FOS_FORCEFILESYSTEM;
            if folders {
                opts |= FOS_PICKFOLDERS;
            }
            if multi {
                opts |= FOS_ALLOWMULTISELECT;
            }
            dlg.SetOptions(opts)?;
            let t = wide(title);
            dlg.SetTitle(PCWSTR(t.as_ptr()))?;
            dlg.Show(Some(owner))?; // Err(ERROR_CANCELLED) when the user cancels
            let results = dlg.GetResults()?;
            (0..results.GetCount()?)
                .map(|i| path_of(&results.GetItemAt(i)?))
                .collect()
        }
    };
    run().unwrap_or_default()
}

pub fn pick_files(owner: HWND) -> Vec<PathBuf> {
    open(owner, "Add files to the workbook", false, true)
}

pub fn pick_folders(owner: HWND) -> Vec<PathBuf> {
    open(owner, "Add folders to the workbook", true, true)
}

pub fn pick_folder(owner: HWND, title: &str) -> Option<PathBuf> {
    open(owner, title, true, false).into_iter().next()
}

/// "Save as" for exporting one file.
pub fn save_file(owner: HWND, suggested_name: &str) -> Option<PathBuf> {
    let run = || -> Result<PathBuf> {
        // SAFETY: as above.
        unsafe {
            let dlg: IFileSaveDialog =
                CoCreateInstance(&FileSaveDialog, None, CLSCTX_INPROC_SERVER)?;
            let base: &IFileDialog = &dlg;
            base.SetOptions(base.GetOptions()? | FOS_FORCEFILESYSTEM)?;
            let n = wide(suggested_name);
            base.SetFileName(PCWSTR(n.as_ptr()))?;
            let t = wide("Export file");
            base.SetTitle(PCWSTR(t.as_ptr()))?;
            base.Show(Some(owner))?;
            path_of(&base.GetResult()?)
        }
    };
    run().ok()
}
