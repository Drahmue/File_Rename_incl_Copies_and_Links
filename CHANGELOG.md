# Changelog

All notable changes to FRICAL are documented here.

## [Unreleased]

## [1.1.0] - 2026-03-12

### Fixed

- **Program freeze on exit**: `show_selection_box` called `root.destroy()` directly inside the button callback while `root.mainloop()` was still running. This left the Tkinter event loop in an undefined state, causing the program to freeze when the user selected "Programm beenden".
  - **Fix**: Replaced `root.destroy()` in the callback with `root.quit()` to stop the mainloop cleanly, then call `root.destroy()` after `mainloop()` returns — consistent with the pattern already used in `show_preview_window`.

- **Tkinter resource leak on repeated runs**: `choose_filename_and_replacementname` and `choose_filename_with_path` created a `tk.Tk()` instance but never called `root.destroy()`, causing orphaned Tk instances to accumulate in memory when the user performed multiple rename operations in one session.
  - **Fix**: Added `root.destroy()` to all exit paths in both functions.

## [1.0.0] - 2025-10-11

### Added

- Initial release of FRICAL v1
- Batch file renaming across recursive directory trees
- Windows shortcut (.lnk) target path updating via COM interface
- Date extraction from filenames (YYYYMMDD) with automatic creation date update
- Preview window showing all files and links to be modified before execution
- Multiprocessing for parallel directory scanning with progress bars
- Reusable file list — scan once, perform multiple rename operations
- Single File Mode: interactive GUI for selecting individual files
- Batch Mode: process multiple rename pairs from a semicolon-separated text file
- Structured logging via `ahlib.ExtendedLogger`
- INI-based configuration via `ahlib.StructuredConfigParser`
- Fallback to single-threaded scanning for small directory trees (< 4 subdirectories)
- `if __name__ == '__main__':` protection to prevent worker process re-execution on Windows
