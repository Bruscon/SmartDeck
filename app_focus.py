from pathlib import Path
import sys
import time
from typing import Optional, List, Tuple
import win32gui
import win32con
import win32api
import win32process
import psutil
import json
from datetime import datetime
import subprocess
import os

class AppFocuser:
    CONFIG_FILE = Path(__file__).parent / "app_focus.json"
    LOG_FILE = Path(__file__).parent / "app_focus.log"
    
    def __init__(self, app_path: str):
        self.app_path = Path(app_path)
        self.app_name = self.app_path.stem.lower()  # Name without extension
        self.load_config()
        
    def load_config(self) -> None:
        """Load or create configuration file with defaults."""
        default_config = {
            "global": {
                "launch_timeout": 10,
                "focus_retry_delay": 0.1,
                "max_retries": 3,
                "max_log_lines": 1000,
                "tab_switch_delay": 0.1,  # Added for consistency
                "max_tabs": 20,  # Added for consistency
            },
            "window_classes": {},  # Map exe names to known window classes
            "last_focused": {},    # Remember last focused window for each app
            "app_configs": {       # Moved from hardcoded APP_CONFIGS
                "bcompare": {
                    "window_classes": ["TViewForm"],
                    "title_required": True,
                    "title_includes": ["Compare"],
                    "title_excludes": ["Home"],
                }
            }
        }
        
        try:
            if self.CONFIG_FILE.exists():
                with open(self.CONFIG_FILE, 'r') as f:
                    self.config = json.load(f)
                    # Ensure all sections exist
                    for key, value in default_config.items():
                        if key not in self.config:
                            self.config[key] = value
                    # Ensure all global parameters exist
                    for key, value in default_config["global"].items():
                        if key not in self.config["global"]:
                            self.config["global"][key] = value
            else:
                self.log("Creating new config file with default settings")
                self.config = default_config
                self.save_config()

        except Exception as e:
            self.log(f"Error handling config file: {e}")
            self.config = default_config

    def log(self, message: str) -> None:
        """Log message with timestamp and app name, maintaining line limit."""
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_message = f'[{timestamp}] [{self.app_name}] {message}\n'
            
            # Read existing logs if file exists
            existing_lines = []
            if self.LOG_FILE.exists():
                with open(self.LOG_FILE, 'r', encoding='utf-8') as f:
                    existing_lines = f.readlines()
            
            # Calculate how many lines to keep
            max_lines = self.config["global"]["max_log_lines"]
            if len(existing_lines) >= max_lines:
                # Keep last (max_lines - 1) lines to make room for new line
                existing_lines = existing_lines[-(max_lines - 1):]
            
            # Write back truncated logs plus new line
            with open(self.LOG_FILE, 'w', encoding='utf-8') as f:
                f.writelines(existing_lines)
                f.write(log_message)
                
        except Exception as e:
            # If logging fails, fail silently but print to stderr for debugging
            print(f"Error writing to log file: {e}", file=sys.stderr)

    def save_config(self) -> None:
        """Save current configuration to file."""
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            self.log(f"Error saving config: {e}")

    def find_app_windows(self) -> List[Tuple[int, int, str]]:
        """Find all windows belonging to the application."""
        app_windows = []
        app_pids = set()  # Track application process IDs
        
        # Special cases configuration - can be moved to config file if needed
        APP_CONFIGS = {
            "bcompare": {
                "window_classes": ["TViewForm"],
                "title_required": True,
                "title_includes": ["Compare"],
                "title_excludes": ["Home"],
            },
            # Add other apps that need special handling here
            # "someapp": {
            #     "window_classes": ["MainWindow", "DocumentWindow"],
            #     "title_required": True,
            #     "title_includes": ["Document"],
            #     "title_excludes": ["Welcome"]
            # }
        }
        
        # Get app-specific config if it exists
        app_config = APP_CONFIGS.get(self.app_name.lower(), {
            "window_classes": [],  # Empty means accept all window classes
            "title_required": False,
            "title_includes": [],
            "title_excludes": []
        })
        
        # First pass: identify all application processes
        def find_processes(hwnd, pids):
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    process = psutil.Process(pid)
                    if process.name().lower() == f"{self.app_name}.exe":
                        pids.add(pid)
                except Exception:
                    pass
            except Exception:
                pass
            return True

        # Second pass: find all windows belonging to application processes
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    # Get window info
                    class_name = win32gui.GetClassName(hwnd)
                    title = win32gui.GetWindowText(hwnd)
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    
                    # Debug: Log window being checked
                    self.log(f"DEBUG: Checking visible window - HWND: {hwnd}, Class: {class_name}, "
                            f"Title: {title}, PID: {pid}")
                    
                    # Check if window belongs to our application
                    if pid in app_pids:
                        # Check window class if specified
                        if (not app_config["window_classes"] or 
                            class_name in app_config["window_classes"]):
                            
                            # Check if we require a title
                            if app_config["title_required"] and not title:
                                return True
                                
                            # Check title includes/excludes
                            if any(exclude in title for exclude in app_config["title_excludes"]):
                                return True
                            if app_config["title_includes"] and not any(
                                include in title for include in app_config["title_includes"]):
                                return True
                                
                            # Window passed all filters
                            self.log(f"DEBUG: Adding window - Class: {class_name}, "
                                    f"Title: {title}, PID: {pid}")
                            windows.append((hwnd, pid, title or f"Window ({class_name})"))
                    
                except Exception as e:
                    self.log(f"Error checking window {hwnd}: {e}")
            return True

        try:
            # First find all application processes
            win32gui.EnumWindows(lambda hwnd, l: find_processes(hwnd, app_pids), None)
            self.log(f"DEBUG: Found application PIDs: {app_pids}")
            
            # Then find all windows belonging to those processes
            win32gui.EnumWindows(lambda hwnd, l: enum_windows_callback(hwnd, app_windows), None)
            
            # Sort windows by handle for consistent ordering
            app_windows.sort(key=lambda x: x[0])
            
            self.log(f"DEBUG: Found total of {len(app_windows)} application windows")
            for hwnd, pid, title in app_windows:
                self.log(f"DEBUG: Final window list - HWND: {hwnd}, PID: {pid}, Title: {title}")
                
        except Exception as e:
            self.log(f"Error enumerating windows: {e}")
            
        return app_windows

    def focus_window(self, hwnd: int) -> bool:
        """Focus the specified window with enhanced focus techniques."""
        max_retries = self.config["global"]["max_retries"]
        retry_delay = self.config["global"]["focus_retry_delay"]
        
        # Get the current foreground window's thread
        cur_foreground = win32gui.GetForegroundWindow()
        cur_thread = win32process.GetWindowThreadProcessId(cur_foreground)[0]
        # Get the target window's thread and process ID
        target_thread, target_pid = win32process.GetWindowThreadProcessId(hwnd)
        
        for attempt in range(max_retries):
            try:
                # Attach input threads
                win32process.AttachThreadInput(target_thread, cur_thread, True)
                try:
                    if win32gui.IsIconic(hwnd):  # If minimized
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    
                    # Force allow setting foreground window (using win32con, not win32gui)
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                    
                    # Save as last focused window
                    self.config["last_focused"][self.app_name] = hwnd
                    self.save_config()
                    return True
                    
                finally:
                    # Always detach threads
                    win32process.AttachThreadInput(target_thread, cur_thread, False)
                    
            except Exception as e:
                if attempt == max_retries - 1:
                    self.log(f"Error focusing window after {max_retries} attempts: {e}")
                    return False
                time.sleep(retry_delay)
                
        return False

    def cycle_app_windows(self) -> bool:
        """Find all app windows and focus the appropriate one."""
        # Get list of all app windows
        windows = self.find_app_windows()
        if not windows:
            self.log("No windows found for the application")
            return False
            
        # Get currently focused window and last focused window for this app
        current_hwnd = win32gui.GetForegroundWindow()
        last_focused_hwnd = self.config["last_focused"].get(self.app_name)
        
        # Create list of window info
        valid_windows = [(hwnd, title) for hwnd, _, title in windows]
        
        self.log(f"Found {len(valid_windows)} windows. Current window: {current_hwnd}, Last focused: {last_focused_hwnd}")
        
        # Log all windows we found
        for hwnd, title in valid_windows:
            self.log(f"Window {hwnd}: '{title}'")
        
        try:
            # First priority: If we have a last focused window and it still exists, use it
            if last_focused_hwnd:
                for hwnd, title in valid_windows:
                    if hwnd == last_focused_hwnd:
                        self.log(f"Focusing last focused window: {hwnd} ({title})")
                        return self.focus_window(hwnd)
            
            # Second priority: If current window is one of our windows, cycle to next
            current_index = -1
            for i, (hwnd, _) in enumerate(valid_windows):
                if hwnd == current_hwnd:
                    current_index = i
                    break
            
            self.log(f"Current window index: {current_index}")
            
            # If current window not in list or no last focused window, proceed with cycling
            next_index = (current_index + 1) % len(valid_windows) if current_index >= 0 else 0
            next_hwnd = valid_windows[next_index][0]
            next_title = valid_windows[next_index][1]
            
            self.log(f"Moving from index {current_index} to {next_index} with title: {next_title}")
            return self.focus_window(next_hwnd)
                
        except Exception as e:
            self.log(f"Error during window cycling: {e}")
            return False

    def launch_app(self) -> bool:
        """Launch the application."""
        try:
            # Ensure parent directories exist in case we need to set working directory
            working_dir = self.app_path.parent
            if not working_dir.exists():
                self.log(f"Working directory does not exist: {working_dir}")
                working_dir = None
                
            subprocess.Popen([str(self.app_path)], cwd=working_dir if working_dir else None)
            self.log(f"Launched application: {self.app_path}")
            return True
            
        except Exception as e:
            self.log(f"Error launching application: {e}")
            return False

    def focus_app(self) -> bool:
        """Main function to find and focus the application."""
        start_time = time.time()
        
        # Check if app is running
        if self.is_process_running():
            self.log("Application is running, attempting to cycle windows")
            if self.cycle_app_windows():
                self.log(f"Successfully focused window in {time.time() - start_time:.2f}s")
                return True
                
        # Launch new instance if not running or no windows found
        self.log("Launching new instance")
        return self.launch_app()

    def is_process_running(self) -> bool:
        """Check if the application is already running."""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower() == f"{self.app_name}.exe":
                    return True
            return False
        except Exception as e:
            self.log(f"Error checking process: {e}")
            return False

def main():
    if len(sys.argv) != 2:
        print("Usage: python app_focus.py <path_to_application>")
        print("Example: python app_focus.py C:\\Program Files\\Beyond Compare 4\\BCompare.exe")
        sys.exit(1)

    app_path = sys.argv[1]
    focuser = AppFocuser(app_path)
    try:
        success = focuser.focus_app()
        sys.exit(0 if success else 1)
    except Exception as e:
        focuser.log(f"Critical error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()