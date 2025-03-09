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
            "app_configs": {       # Configurations for special apps
                "bcompare": {
                    "window_classes": ["TViewForm"],
                    "title_required": True,
                    "title_includes": ["Compare"],
                    "title_excludes": ["Home"],
                },
                "git-bash": {
                    "window_classes": ["mintty"],
                    "title_required": True,
                    "title_includes": ["MINGW64"],
                    "title_excludes": [],
                    "process_name": "mintty.exe"  # Special case for Git Bash
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
                    
                    # Ensure all app configs are present
                    if "app_configs" not in self.config:
                        self.config["app_configs"] = {}
                    
                    # Add git-bash config if it doesn't exist
                    if "git-bash" not in self.config["app_configs"]:
                        self.config["app_configs"]["git-bash"] = default_config["app_configs"]["git-bash"]
                        self.log("Added Git Bash configuration to config file")
                        self.save_config()
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
        
        # Get app-specific config if it exists
        app_config = self.config["app_configs"].get(self.app_name.lower(), {
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
                    
                    # Check if this is a special case app (like git-bash using mintty)
                    process_name = app_config.get("process_name", f"{self.app_name}.exe").lower()
                    
                    if process.name().lower() == process_name:
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
                    
                    # Special handling for git-bash/mintty
                    if self.app_name.lower() == "git-bash" and class_name == "mintty":
                        if "MINGW64" in title:
                            self.log(f"DEBUG: Adding window - Class: {class_name}, "
                                    f"Title: {title}, PID: {pid}")
                            windows.append((hwnd, pid, title or f"Window ({class_name})"))
                            return True
                    
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
        
        for attempt in range(max_retries):
            try:
                # Check if window still exists
                if not win32gui.IsWindow(hwnd):
                    return False
                    
                # Simple focus attempt first
                if win32gui.IsIconic(hwnd):  # If minimized
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                    # Save as last focused window only on success
                    self.config["last_focused"][self.app_name] = hwnd
                    self.save_config()
                    return True
                except Exception:
                    # If simple focus fails, try the thread attachment method
                    cur_foreground = win32gui.GetForegroundWindow()
                    if cur_foreground and win32gui.IsWindow(cur_foreground):
                        cur_thread = win32process.GetWindowThreadProcessId(cur_foreground)[0]
                        target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
                        
                        if cur_thread and target_thread and cur_thread != target_thread:
                            win32process.AttachThreadInput(target_thread, cur_thread, True)
                            try:
                                win32gui.SetForegroundWindow(hwnd)
                                win32gui.BringWindowToTop(hwnd)
                                self.config["last_focused"][self.app_name] = hwnd
                                self.save_config()
                                return True
                            finally:
                                win32process.AttachThreadInput(target_thread, cur_thread, False)
                    
            except Exception as e:
                if attempt == max_retries - 1:
                    self.log(f"Error focusing window after {max_retries} attempts: {e}")
                    # Don't return False yet - let cycle_app_windows try other windows
                time.sleep(retry_delay)
                
        return False

    def get_window_z_order(self, hwnd: int) -> int:
        """Get the Z-order index of a window. Lower numbers are more recent."""
        try:
            z_order = 0
            current = win32gui.GetWindow(win32gui.GetDesktopWindow(), win32con.GW_CHILD)
            
            while current:
                if current == hwnd:
                    return z_order
                z_order += 1
                current = win32gui.GetWindow(current, win32con.GW_HWNDNEXT)
                
            return float('inf')  # Window not found
        except Exception as e:
            self.log(f"Error getting Z-order for window {hwnd}: {e}")
            return float('inf')

    def cycle_app_windows(self) -> bool:
        """Find all app windows and focus the appropriate one."""
        windows = self.find_app_windows()
        if not windows:
            self.log("No windows found for the application")
            return False
            
        # Convert to list of (hwnd, pid, title, z_order)
        window_info = []
        for hwnd, pid, title in windows:
            z_order = self.get_window_z_order(hwnd)
            window_info.append((hwnd, pid, title, z_order))
        
        self.log(f"Found {len(window_info)} windows:")
        for hwnd, pid, title, z_order in window_info:
            self.log(f"Window {hwnd} (Z-order: {z_order}): '{title}'")
        
        try:
            current_hwnd = win32gui.GetForegroundWindow()
            
            # If no window is currently focused, start with most recent
            if current_hwnd not in [w[0] for w in window_info]:
                window_info.sort(key=lambda x: x[3])  # Sort by Z-order
                self.log(f"No window focused, starting with most recent: {window_info[0][0]}")
                return self.focus_window(window_info[0][0])
                
            # Otherwise cycle through all windows
            current_index = -1
            for i, (hwnd, _, _, _) in enumerate(window_info):
                if hwnd == current_hwnd:
                    current_index = i
                    break
                    
            next_index = (current_index + 1) % len(window_info)
            next_window = window_info[next_index]
            
            self.log(f"Cycling from index {current_index} to {next_index}: {next_window[0]} ({next_window[2]})")
            return self.focus_window(next_window[0])
                
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
        if self.is_control_pressed():  # Check if Ctrl is pressed
            self.log("Ctrl key detected - launching new instance")
            return self.launch_app()
        
        # logic for cycling windows
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
            app_config = self.config["app_configs"].get(self.app_name.lower(), {})
            process_name = app_config.get("process_name", f"{self.app_name}.exe").lower()
            
            # Special case for git-bash
            if self.app_name.lower() == "git-bash":
                # Check for mintty processes with MINGW64 windows
                for hwnd in self.get_matching_windows("mintty", "MINGW64"):
                    return True
            
            # Regular process check
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower() == process_name:
                    return True
            
            return False
        except Exception as e:
            self.log(f"Error checking process: {e}")
            return False
            
    def get_matching_windows(self, class_name, title_substring):
        """Find windows matching the given class name and title substring."""
        matching_windows = []
        
        def enum_callback(hwnd, result_list):
            if win32gui.IsWindowVisible(hwnd):
                window_class = win32gui.GetClassName(hwnd)
                window_title = win32gui.GetWindowText(hwnd)
                
                if window_class == class_name and title_substring in window_title:
                    result_list.append(hwnd)
            return True
            
        win32gui.EnumWindows(lambda hwnd, l: enum_callback(hwnd, matching_windows), None)
        return matching_windows

    def is_control_pressed(self):
        return win32api.GetAsyncKeyState(win32con.VK_CONTROL) & 0x8000 != 0

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