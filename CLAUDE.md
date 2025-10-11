# FRICAL - File Rename Including Copies and Links

## Project Overview

FRICAL is a Windows-based Python tool designed to batch rename files along with all their copies and associated Windows shortcuts (.lnk files) across a directory structure. The tool features multiprocessing for fast directory scanning, a preview window for review before execution, and structured logging.

## Key Features

- **Batch File Renaming**: Rename original files and all copies found recursively in directory trees
- **Shortcut Management**: Automatically updates target paths in Windows .lnk files
- **Date Extraction**: Extracts YYYYMMDD dates from filenames and updates file creation dates
- **Preview Window**: Shows all files and links to be modified before execution with confirm/cancel options
- **Multiprocessing**: Parallel directory scanning for fast processing of large directory structures
- **Progress Bars**: Real-time progress feedback during file list creation
- **Reusable File List**: Create file list once and perform multiple rename operations without rescanning
- **Structured Logging**: Extended logging to file and console using ahlib.ExtendedLogger
- **Dual Operation Modes**:
  - Single File Mode: Interactive GUI for selecting individual files
  - Batch Mode: Process multiple files from a semicolon-separated list

## Architecture

### Core Functions

- `create_file_list(Such_Pfad)` - Recursively scans directory tree using multiprocessing; returns all file paths
- `_scan_directory_recursive(directory_path)` - Worker function for parallel directory scanning
- `Search_Files(Search, all_files)` - Filters file list by name pattern
- `Rename_Files_and_new_date(Liste, NewName)` - Renames files and updates creation dates based on extracted dates
- `Rename_Files(Liste, NewName)` - Renames files without date adjustment
- `Target_in_LNK(Liste, New_Target)` - Updates Windows shortcut target paths using COM interface
- `Date_Extract(string)` - Extracts dates in YYYYMMDD format from strings
- `set_creation_date(file_path, new_creation_date)` - Sets Windows file creation timestamp
- `load_2column_list(file_path)` - Loads batch rename pairs from text file
- `show_preview_window(old_name, new_name, files_to_rename, links_to_adjust)` - Displays preview with confirmation dialog
- `perform_rename_operation(all_files_list, Basis_Pfad, logger_instance)` - Main operation handler for single rename cycle

### Configuration

The tool uses INI-based configuration managed by `ahlib.StructuredConfigParser`:

**Configuration File**: `FRICAL_v1.ini`

```ini
[Files]
Such_Pfad = \\WIN-H7BKO5H0RMC\Dataserver
Basis_Pfad = \\WIN-H7BKO5H0RMC\Dataserver\Korrespondenz\Post Archiv
logfile = \\WIN-H7BKO5H0RMC\Dataserver\Python Hilfsdateien\Email_Analyse_Log.txt
```

- **Such_Pfad**: Root directory for recursive file search
- **Basis_Pfad**: Base directory for file operations
- **logfile**: Path to log file (created automatically if parent directory exists)

### Multiprocessing Architecture

The tool uses `concurrent.futures.ProcessPoolExecutor` for parallel directory scanning:

1. **Directory Splitting**: Splits search path at first level into subdirectories
2. **Worker Pool**: Creates worker processes (up to CPU count)
3. **Parallel Scanning**: Each worker scans one subdirectory recursively
4. **Result Aggregation**: Collects results with progress bar
5. **Fallback Logic**: Falls back to single-threaded mode for small directory trees (< 4 subdirectories)

**Important**: All initialization code is protected by `if __name__ == '__main__':` to prevent worker processes from re-executing configuration and logging setup.

## Dependencies

- `pywin32>=305` - Windows COM interface for shortcut manipulation
- `tqdm>=4.65.0` - Progress bar display
- `ahlib` - Structured configuration parsing and extended logging (installed from GitHub)
- `tkinter` - GUI file dialogs (included with Python)
- `pathlib` - Modern path handling
- `concurrent.futures` - Multiprocessing support (standard library)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/Drahmue/File_Rename_incl_Copies_and_Links.git
cd File_Rename_incl_Copies_and_Links
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Create configuration file `FRICAL_v1.ini` in the same directory as the script:
```ini
[Files]
Such_Pfad = <your search path>
Basis_Pfad = <your base path>
logfile = <path to log file>
```

## Usage

### Workflow Overview

1. **Startup**: Script creates complete file list once using multiprocessing (with progress bar)
2. **Operation Loop**: Perform rename operations repeatedly using the same file list
3. **Preview**: Review all changes before execution
4. **Continue/Exit**: Choose to perform another operation or exit

### Single File Mode

1. Run the script: `python FRICAL_v1.py`
2. Wait for file list creation (progress bar shows status)
3. Select "Einzelfile" option
4. Choose file to rename via GUI dialog
5. Enter new filename
6. **Review preview window** showing:
   - Old and new filenames
   - All files to be renamed (with full paths)
   - All .lnk files to be adjusted (with full paths)
   - Summary count
7. Click green "Umbenennung durchführen" to proceed or red "Abbrechen" to cancel
8. Script processes the file, all copies, and associated .lnk files
9. Choose "Weiteren Vorgang durchführen" to rename more files or "Programm beenden" to exit

### Batch Mode

1. Create a text file with old;new filename pairs (one per line)
2. Run the script: `python FRICAL_v1.py`
3. Wait for file list creation
4. Select "Batchdatei" option
5. Choose the batch file
6. For each entry, review preview window and confirm or cancel
7. Script processes all confirmed files in the list
8. Choose to continue with another operation or exit

## Batch File Format

```
oldfilename.ext;newfilename.ext
another_old.pdf;another_new.pdf
# Comments start with #
```

- Each line contains old and new filename separated by semicolon
- Include file extensions
- No path information (filenames only)
- Lines starting with # are ignored
- No empty lines at end of file

## Current Status

**Version**: v1 (FRICAL_v1.py)
**Last Updated**: 2025-01-11

### Recent Changes

- **2025-01-11**: Updated .gitignore to exclude _Archiv/ folder
- **2025-01-11**: Updated CLAUDE.md and requirements.txt documentation
- **2025-10-11**: Fixed preview window event handling (button callbacks now properly trigger)
- **2025-10-11**: Added multiprocessing for fast directory scanning with progress bars
- **2025-10-11**: Added preview window with file/link listing before execution
- **2025-10-11**: Implemented reusable file list with continue/exit option
- **2025-10-11**: Migrated to ahlib for configuration and logging
- **2025-10-11**: Fixed multiprocessing duplicate output issue with `if __name__ == '__main__':` protection
- **2025-10-11**: Initial GitHub repository setup

## Known Limitations

- Does not find .lnk files with different names than their targets
- Windows-specific due to .lnk file handling
- Requires network path access for default configuration
- File list is created once at startup and not refreshed during execution
- Multiprocessing requires proper `if __name__ == '__main__':` protection on Windows

## Technical Notes

### Preview Window Implementation

The preview window uses tkinter with a dictionary-based result storage pattern:
- Creates `result = {'confirmed': False}` outside Tkinter context
- Callbacks modify the dictionary value
- Uses `root.quit()` followed by `root.destroy()` for clean shutdown
- Returns `result['confirmed']` which remains accessible after window destruction

### Multiprocessing Pattern

To prevent worker processes from re-executing initialization code:
- All config loading, logger creation, and main loop are inside `if __name__ == '__main__':` block
- Worker processes only import function definitions
- Logger is passed as parameter to functions that need it
- Global logger variable is set via `global logger` declaration in functions

## Future Enhancements

- Add error handling for file access issues during rename operations
- Implement dry-run mode for testing
- Add option to refresh file list during execution
- Cross-platform support for non-.lnk shortcuts
- Better handling of file conflicts
- Add undo functionality
- Support for regex-based rename patterns
