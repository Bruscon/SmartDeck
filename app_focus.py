from pathlib import Path
import sys
import time
from typing import Optional, List, Tuple, Dict, Set
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
    
    def __init__(self, app_path: str, debug: bool = True):
        self.app_path = Path(app_path)
        self.app_name = self.app_path.stem.lower()  # Name without extension
        self.debug = debug
        self.load_config()
        
    def load_config(self) -> None:
        """Load or create configuration file with defaults."""
        default_config = {
            "global": {
                "launch_timeout": 5,  # Reduced from 10
                "focus_retry_delay": 0.02,  # Reduced from 0.1
                "max_retries": 3,
                "max_log_lines": 1000,
                "tab_switch_delay": 0.02,  # Reduced from 0.1
                "max_tabs": 20,
                "debug": False
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
        """Log message with timestamp and app name, ensuring it's written to file."""
        try:
            # Format with timestamp and app name
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_message = f'[{timestamp}] [{self.app_name}] {message}'
            
            # Always print to console (this helps with immediate feedback)
            print(log_message)
            
            # Determine log file path
            log_file = self.LOG_FILE
            
            # Make sure the log directory exists
            log_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Append to the log file
            with open(log_file, 'a', encoding='utf-8', errors='replace') as f:
                f.write(log_message + '\n')
                f.flush()  # Ensure it's written immediately
                
        except Exception as e:
            # If logging fails, print to stderr
            print(f"Error writing to log file: {e}", file=sys.stderr)

    def save_config(self) -> None:
        """Save current configuration to file."""
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            self.log(f"Error saving config: {e}")

    # Insert this into app_focus.py to replace the find_app_windows method

    def find_app_windows(self) -> List[Tuple[int, int, str]]:
        """Find all windows belonging to the application with optimized enumeration."""
        app_windows = []
        all_windows = []  # Cache all windows in one enumeration
        
        # Get app-specific config if it exists
        app_config = self.config["app_configs"].get(self.app_name.lower(), {
            "window_classes": [],  # Empty means accept all window classes
            "title_required": False,
            "title_includes": [],
            "title_excludes": []
        })
        
        # First, identify all matching processes
        app_pids = set()
        process_name = app_config.get("process_name", f"{self.app_name}.exe").lower()
        
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_name = proc.info['name'].lower()
                if proc_name == process_name or self.app_name.lower() in proc_name:
                    app_pids.add(proc.info['pid'])
                    self.log(f"Found matching process: {proc_name} (PID: {proc.info['pid']})")
            except Exception as e:
                self.log(f"Error checking process: {e}")
        
        # Now enumerate all windows - this time including ALL windows regardless of visibility
        def enum_all_windows(hwnd, results):
            try:
                class_name = win32gui.GetClassName(hwnd)
                title = win32gui.GetWindowText(hwnd)
                
                # Skip known system window classes that aren't application windows
                if class_name in ["ToolTip", "SysListView32", "Button", "Static", 
                                 "DummyDWMListenerWindow", "ApplicationFrameWindow"]:
                    return True
                
                # Get basic process information for this window
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    results.append((hwnd, class_name, title, pid))
                except Exception:
                    # If we can't get process info, still track the window for debugging
                    results.append((hwnd, class_name, title, None))
                    
            except Exception:
                pass
            return True
        
        # Collect all windows in one pass
        win32gui.EnumWindows(lambda hwnd, l: enum_all_windows(hwnd, all_windows), None)
        
        self.log(f"Found {len(all_windows)} windows to check")
        
        # Map PIDs to executable names for better matching
        pid_to_exe = {}
        for pid in app_pids:
            try:
                proc = psutil.Process(pid)
                pid_to_exe[pid] = proc.name().lower()
            except Exception:
                pass
        
        # Process all windows
        for hwnd, class_name, title, pid in all_windows:
            matches_app = False
            match_reason = []
            
            # 1. Check if window belongs to one of our known app processes
            if pid in app_pids:
                matches_app = True
                match_reason.append(f"PID match ({pid})")
            
            # 2. Check window class (if specified)
            if app_config.get("window_classes") and class_name in app_config.get("window_classes", []):
                matches_app = True
                match_reason.append(f"Class match ({class_name})")
            
            # 3. Check window title
            title_includes = app_config.get("title_includes", [])
            if title_includes and any(include in title for include in title_includes):
                matches_app = True
                match_reason.append("Title include match")
            
            # 4. If the app name is in the window title, it's likely a match
            if self.app_name.lower() in title.lower():
                matches_app = True
                match_reason.append("App name in title")
            
            # Skip if no match found
            if not matches_app:
                continue
            
            # Apply filters if this window matches our app
            
            # Filter 1: Title requirements
            if app_config.get("title_required", False) and not title:
                self.log(f"Skipping window {hwnd} - no title but title required")
                continue
            
            # Filter 2: Title exclusions
            if any(exclude in title for exclude in app_config.get("title_excludes", [])):
                self.log(f"Skipping window {hwnd} - matched title exclusion")
                continue
            
            # Filter 3: Add the window to our list
            self.log(f"Adding window - HWND: {hwnd}, Title: '{title}', Class: {class_name}, Reason: {', '.join(match_reason)}")
            app_windows.append((hwnd, pid if pid else 0, title or f"Window ({class_name})"))
        
        # Return the list of matching windows
        self.log(f"Found {len(app_windows)} matching windows for {self.app_name}")
        for hwnd, pid, title in app_windows:
            self.log(f"Window: {hwnd} '{title}' (PID: {pid})")
            
        return app_windows

    def focus_window(self, hwnd: int) -> bool:
        """Focus a specific window with enhanced techniques that work across virtual desktops."""
        try:
            # Check if window exists
            if not win32gui.IsWindow(hwnd):
                self.log(f"Window {hwnd} does not exist")
                return False
            
            # Get window info for logging
            try:
                window_title = win32gui.GetWindowText(hwnd)
                self.log(f"Attempting to focus window {hwnd}: '{window_title}'")
            except Exception:
                self.log(f"Attempting to focus window {hwnd}")
            
            # If minimized, restore it
            if win32gui.IsIconic(hwnd):
                self.log(f"Window is minimized, restoring")
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.05)  # Short delay to let restore complete
            
            # TECHNIQUE 1: Simple focus attempt
            try:
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                
                # Check if focus succeeded
                new_foreground = win32gui.GetForegroundWindow()
                if new_foreground == hwnd:
                    self.log("Simple focus succeeded")
                    # Save as last focused window on success
                    self.config["last_focused"][self.app_name] = hwnd
                    self.save_config()
                    return True
            except Exception as e:
                self.log(f"Simple focus failed: {e}")
            
            # TECHNIQUE 2: ALT key trick to bypass Windows restrictions
            try:
                # Simulate Alt keypress - this often helps bypass Windows restrictions
                win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
                time.sleep(0.05)
                
                # Try focus while ALT is pressed
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                
                # Release Alt key
                win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.05)
                
                # Check if focus succeeded
                new_foreground = win32gui.GetForegroundWindow()
                if new_foreground == hwnd:
                    self.log("ALT key focus succeeded")
                    # Save as last focused window on success
                    self.config["last_focused"][self.app_name] = hwnd
                    self.save_config()
                    return True
            except Exception as e:
                self.log(f"ALT key focus failed: {e}")
            
            # TECHNIQUE 3: Thread attachment approach
            try:
                # Get current foreground window
                current_foreground = win32gui.GetForegroundWindow()
                if current_foreground and win32gui.IsWindow(current_foreground):
                    # Get thread IDs
                    cur_thread = win32process.GetWindowThreadProcessId(current_foreground)[0]
                    target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
                    
                    if cur_thread and target_thread and cur_thread != target_thread:
                        self.log(f"Attempting thread attachment between {cur_thread} and {target_thread}")
                        
                        # Attach input processing
                        win32process.AttachThreadInput(target_thread, cur_thread, True)
                        try:
                            # Multiple focus attempts with different APIs
                            win32gui.BringWindowToTop(hwnd)
                            win32gui.SetForegroundWindow(hwnd)
                            win32gui.SetActiveWindow(hwnd)
                            time.sleep(0.05)
                            
                            # Verify focus success
                            new_foreground = win32gui.GetForegroundWindow()
                            if new_foreground == hwnd:
                                self.log("Thread attachment focus succeeded")
                                # Save as last focused window on success
                                self.config["last_focused"][self.app_name] = hwnd
                                self.save_config()
                                return True
                        finally:
                            # Always detach threads
                            win32process.AttachThreadInput(target_thread, cur_thread, False)
            except Exception as e:
                self.log(f"Thread attachment focus failed: {e}")
            
            # TECHNIQUE 4: AllowSetForegroundWindow approach
            try:
                # Get target window's process ID
                target_pid = win32process.GetWindowThreadProcessId(hwnd)[1]
                
                # Allow that process to set foreground window
                import win32security
                win32security.AllowSetForegroundWindow(target_pid)
                
                # Multiple focus attempts
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                
                # Check result
                new_foreground = win32gui.GetForegroundWindow()
                if new_foreground == hwnd:
                    self.log("AllowSetForegroundWindow focus succeeded")
                    # Save as last focused window on success
                    self.config["last_focused"][self.app_name] = hwnd
                    self.save_config()
                    return True
            except Exception as e:
                self.log(f"AllowSetForegroundWindow focus failed: {e}")
            
            # All techniques failed
            self.log("All focus techniques failed")
            return False
        except Exception as e:
            self.log(f"Error in focus_window: {e}")
            return False

    def get_window_z_order(self, hwnd: int) -> int:
        """Get the Z-order index of a window. Lower numbers are more recent."""
        try:
            z_order = 0
            current = win32gui.GetWindow(win32gui.GetDesktopWindow(), win32con.GW_CHILD)
            
            while current and z_order < 100:  # Limit search depth
                if current == hwnd:
                    return z_order
                z_order += 1
                current = win32gui.GetWindow(current, win32con.GW_HWNDNEXT)
                
            return float('inf')  # Window not found
        except Exception as e:
            self.log(f"Error getting Z-order for window {hwnd}: {e}")
            return float('inf')

    def cycle_app_windows(self) -> bool:
        """Find all app windows and focus the appropriate one with improved reliability."""
        windows = self.find_app_windows()
        if not windows:
            self.log("No windows found for the application")
            return False
        
        self.log(f"Found {len(windows)} windows for cycling")
        
        try:
            # First, check if we have a valid last focused window
            last_hwnd = self.config["last_focused"].get(self.app_name)
            if last_hwnd:
                # Verify this window is still valid
                window_exists = False
                for hwnd, _, title in windows:
                    if hwnd == last_hwnd:
                        window_exists = True
                        self.log(f"Found last focused window: {hwnd} '{title}'")
                        break
                
                if not window_exists:
                    self.log(f"Last focused window {last_hwnd} no longer exists")
                    # Remove invalid handle from config
                    if self.app_name in self.config["last_focused"]:
                        del self.config["last_focused"][self.app_name]
                        self.save_config()
            
            # Get current foreground window
            current_hwnd = win32gui.GetForegroundWindow()
            
            # Special case for Sublime Text: If only one window is found, always focus it
            if len(windows) == 1 and self.app_name.lower() == "sublime_text":
                self.log(f"Only one Sublime Text window found, focusing it: {windows[0][0]}")
                return self.focus_window(windows[0][0])
            
            # Check if one of our windows is currently focused
            current_window_index = -1
            for i, (hwnd, _, title) in enumerate(windows):
                if hwnd == current_hwnd:
                    current_window_index = i
                    self.log(f"Currently focused window index {i}: {hwnd} '{title}'")
                    break
            
            # If current window is not in our list, or we only have one window, focus first window
            if current_window_index == -1 or len(windows) == 1:
                self.log(f"No application window currently focused, focusing first window: {windows[0][0]}")
                return self.focus_window(windows[0][0])
            
            # If we have multiple windows, cycle to next one
            next_index = (current_window_index + 1) % len(windows)
            next_window = windows[next_index]
            
            self.log(f"Cycling from index {current_window_index} to {next_index}: {next_window[0]} '{next_window[2]}'")
            focus_success = self.focus_window(next_window[0])
            
            # If focusing failed, try another window
            if not focus_success and len(windows) > 1:
                self.log("Focus failed, trying another window")
                next_next_index = (next_index + 1) % len(windows)
                next_next_window = windows[next_next_index]
                self.log(f"Trying window at index {next_next_index}: {next_next_window[0]} '{next_next_window[2]}'")
                return self.focus_window(next_next_window[0])
                
            return focus_success
                
        except Exception as e:
            self.log(f"Error during window cycling: {e}")
            
            # On error, try the first window as a fallback
            if windows:
                self.log("Trying first window as fallback")
                return self.focus_window(windows[0][0])
            
            return False

    def launch_app(self) -> bool:
        """Launch the application."""
        try:
            # Set high priority for the launched process
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = win32con.SW_SHOWNORMAL
            
            # Ensure parent directories exist in case we need to set working directory
            working_dir = self.app_path.parent
            if not working_dir.exists():
                self.log(f"Working directory does not exist: {working_dir}")
                working_dir = None
            
            # Create process with high priority
            process = subprocess.Popen(
                [str(self.app_path)], 
                cwd=working_dir if working_dir else None,
                startupinfo=startupinfo
            )
            
            # Set the process priority
            try:
                handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, process.pid)
                win32process.SetPriorityClass(handle, win32process.ABOVE_NORMAL_PRIORITY_CLASS)
            except Exception as e:
                self.log(f"Warning: Could not set process priority: {e}")
            
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
        
        # Find all windows first to have a complete picture
        windows = self.find_app_windows()
        if not windows:
            self.log("No windows found for the application, launching new instance")
            return self.launch_app()
            
        self.log(f"Found {len(windows)} windows for application")
        
        # Get current foreground window to see if it's one of our app windows
        current_hwnd = win32gui.GetForegroundWindow()
        
        # Check if last focused window is valid and is NOT currently focused
        last_hwnd = self.config["last_focused"].get(self.app_name)
        last_window_valid = False
        
        if last_hwnd:
            if not win32gui.IsWindow(last_hwnd):
                self.log(f"Last focused window {last_hwnd} no longer exists")
                # Remove invalid handle from config
                if self.app_name in self.config["last_focused"]:
                    del self.config["last_focused"][self.app_name]
                    self.save_config()
            else:
                last_window_valid = True
                self.log(f"Last focused window {last_hwnd} is valid")
                
                # Check if this window is already in focus
                if current_hwnd == last_hwnd:
                    self.log("Last focused window is currently in focus - need to cycle")
                    # Find position of current window in our list to cycle to next
                    current_index = -1
                    for i, (hwnd, _, title) in enumerate(windows):
                        if hwnd == current_hwnd:
                            current_index = i
                            break
                            
                    if current_index != -1 and len(windows) > 1:
                        # Cycle to next window
                        next_index = (current_index + 1) % len(windows)
                        next_hwnd = windows[next_index][0]
                        self.log(f"Cycling from window {current_index} to {next_index}: {next_hwnd}")
                        return self.focus_window(next_hwnd)
                    
        # Check if any of our app windows is currently focused
        is_app_window_focused = False
        focused_window_index = -1
        
        for i, (hwnd, _, title) in enumerate(windows):
            if hwnd == current_hwnd:
                is_app_window_focused = True
                focused_window_index = i
                self.log(f"App window already in focus: {hwnd} '{title}'")
                break
        
        # If an app window is focused, cycle to the next one
        if is_app_window_focused and len(windows) > 1:
            next_index = (focused_window_index + 1) % len(windows)
            next_hwnd = windows[next_index][0]
            self.log(f"Cycling from window {focused_window_index} to {next_index}: {next_hwnd}")
            return self.focus_window(next_hwnd)
        
        # If no app window is focused but we have a valid last window, focus that
        if last_window_valid:
            self.log(f"Focusing last used window: {last_hwnd}")
            return self.focus_window(last_hwnd)
        
        # Otherwise focus the first available window
        self.log(f"Focusing first available window: {windows[0][0]}")
        return self.focus_window(windows[0][0])

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
            
            # Optimized process check - terminate early on first match
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

def set_process_priority():
    """Set the current process to high priority."""
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, 
                                     win32api.GetCurrentProcessId())
        win32process.SetPriorityClass(handle, win32process.HIGH_PRIORITY_CLASS)
        return True
    except Exception:
        return False

def main():
    # Set script to high priority immediately
    set_process_priority()
    
    if len(sys.argv) < 2:
        print("Usage: python app_focus.py <path_to_application> [--debug]")
        print("Example: python app_focus.py C:\\Program Files\\Beyond Compare 4\\BCompare.exe")
        sys.exit(1)

    # Check for debug flag
    debug_mode = "--debug" in sys.argv
    
    # Get app path (the first non-flag argument)
    app_path = next((arg for arg in sys.argv[1:] if not arg.startswith("--")), None)
    if not app_path:
        print("Error: No application path specified")
        sys.exit(1)

    focuser = AppFocuser(app_path, debug=debug_mode)
    try:
        success = focuser.focus_app()
        sys.exit(0 if success else 1)
    except Exception as e:
        focuser.log(f"Critical error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()