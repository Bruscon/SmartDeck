from pathlib import Path
import sys
import time
from typing import Optional, Dict, Any
import win32gui
import win32con
import win32api
import win32process
import keyboard
import json
from datetime import datetime

class ChromeTabSwitcher:
    CHROME_CLASS = "Chrome_WidgetWin_1"
    CONFIG_FILE = Path(__file__).parent / "chrome_tab_switcher.json"
    LOG_FILE = Path(__file__).parent / "chrome_tab_switcher.log"
    
    def __init__(self, domain: str):
        self.domain = domain.lower()  # Normalize domain to lowercase
        if self.domain.startswith("www."):
            self.domain = self.domain[4:]  # Remove www. prefix if present
        self.load_config()
        self._set_process_priority()

    def load_config(self) -> None:
        """Load or create configuration file with defaults."""
        default_config = {
            "global": {
                "max_tabs": 20,
                "tab_switch_delay": 0.1,
                "launch_timeout": 10,
                "focus_retry_delay": 0.1,
                "max_retries": 3,
                "max_log_lines": 1000
            },
            "last_positions": {},
            "urls": {
                "claude.ai": "https://claude.ai",
                "perplexity.ai": "https://www.perplexity.ai"
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

            # Ensure current domain exists in config
            if self.domain not in self.config["urls"]:
                self.config["urls"][self.domain] = f"https://{self.domain}"
                self.save_config()
            if self.domain not in self.config["last_positions"]:
                self.config["last_positions"][self.domain] = 0
                self.save_config()

        except Exception as e:
            self.log(f"Error handling config file: {e}")
            self.config = default_config

    def log(self, message: str) -> None:
        """Log message with timestamp and domain, maintaining line limit."""
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_message = f'[{timestamp}] [{self.domain}] {message}\n'
            
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

    def _set_process_priority(self) -> None:
        """Set script to high priority for faster tab switching."""
        try:
            process = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, 
                                         win32api.GetCurrentProcessId())
            win32process.SetPriorityClass(process, win32process.HIGH_PRIORITY_CLASS)
        except Exception as e:
            self.log(f"Error setting process priority: {e}")

    def find_chrome_window(self) -> Optional[int]:
        """Find Chrome window with optimized search."""
        try:
            # Fast direct window search first
            hwnd = win32gui.FindWindow(self.CHROME_CLASS, None)
            if hwnd and win32gui.IsWindowVisible(hwnd):
                return hwnd

            # Fallback to window enumeration
            def enum_windows_callback(hwnd, chrome_windows):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == self.CHROME_CLASS:
                    chrome_windows.append(hwnd)
                return True

            chrome_windows = []
            win32gui.EnumWindows(lambda hwnd, l: enum_windows_callback(hwnd, chrome_windows), None)
            
            return chrome_windows[0] if chrome_windows else None

        except Exception as e:
            self.log(f"Error finding Chrome window: {e}")
            return None

    def launch_chrome(self, new_tab: bool = False) -> bool:
        """Launch Chrome with the domain URL."""
        import os
        import subprocess
        
        try:
            if not self.domain:
                url = "chrome://newtab"  # Just open Chrome
            else:
                url = self.config["urls"][self.domain]
            
            if new_tab:
                url = "--new-tab " + url
            # Check multiple possible Chrome locations
            chrome_paths = [
                Path(os.environ['ProgramFiles(x86)']) / 'Google/Chrome/Application/chrome.exe',
                Path(os.environ['ProgramFiles']) / 'Google/Chrome/Application/chrome.exe',
                Path(os.environ['LocalAppData']) / 'Google/Chrome/Application/chrome.exe',
            ]
            
            for path in chrome_paths:
                if path.exists():
                    self.log(f"Launching Chrome from: {path}")
                    subprocess.Popen([str(path), url])
                    return True
            
            # Fallback to PATH
            subprocess.Popen(['chrome', url])
            self.log("Launched Chrome using PATH")
            return True
            
        except Exception as e:
            self.log(f"Failed to launch Chrome: {e}")
            return False

    def focus_window(self, hwnd: int) -> bool:
        """Focus window with retry logic."""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return True
            except Exception as e:
                if attempt == max_retries - 1:
                    self.log(f"Error focusing window after {max_retries} attempts: {e}")
                    return False
                time.sleep(0.1)  # Brief pause before retry

    def cycle_through_tabs(self) -> bool:
        """Cycle through tabs with thorough tab enumeration."""
        if not self.domain:  # No domain means just focus Chrome
            return False
            
        # Check current tab first
        hwnd = win32gui.GetForegroundWindow()
        current_title = win32gui.GetWindowText(hwnd)
        self.log(f"\n=== Starting tab cycle ===")
        self.log(f"Starting with window title: {current_title}")
        
        if self.is_matching_tab(current_title):
            self.log("Current tab matches target domain")
            return True

        # Start by going to the leftmost tab
        keyboard.press_and_release('ctrl+1')
        time.sleep(self.config["global"]["tab_switch_delay"] * 2)  # Extra delay for stability
        
        max_tabs = self.config["global"]["max_tabs"]
        tab_delay = self.config["global"]["tab_switch_delay"]
        
        # Keep track of titles we've seen to avoid infinite loops
        seen_titles = {}  # Use dict to track position of each title
        tabs_checked = 0
        
        # First check leftmost tab
        first_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        self.log(f"\nStarting from leftmost tab ({tabs_checked + 1}): {first_title}")
        
        if self.is_matching_tab(first_title):
            self.log("Found matching tab at first position")
            self.config["last_positions"][self.domain] = 0
            self.save_config()
            return True
            
        seen_titles[first_title] = 0
        tabs_checked += 1
        
        # Cycle through remaining tabs
        while tabs_checked < max_tabs:
            # Use ctrl+tab to move right one tab
            keyboard.press_and_release('ctrl+tab')
            time.sleep(tab_delay * 1.5)  # Slightly longer delay for stability
            
            current_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            self.log(f"\nChecking tab {tabs_checked + 1}: {current_title}")
            
            if current_title in seen_titles:
                self.log(f"Completed full cycle through {tabs_checked} tabs")
                break
                
            seen_titles[current_title] = tabs_checked
            
            # Debug output current URL patterns
            if 'perplexity' in current_title.lower():
                self.log(f"Found potential perplexity tab with title: {current_title}")
            
            if self.is_matching_tab(current_title):
                self.log(f"Found matching tab at position {tabs_checked}")
                self.config["last_positions"][self.domain] = tabs_checked
                self.save_config()
                return True
                
            tabs_checked += 1
        
        self.log(f"\n=== Tab cycle summary ===")
        self.log(f"Checked {tabs_checked} tabs total")
        self.log("All tabs found:")
        for title, pos in seen_titles.items():
            self.log(f"Position {pos}: {title}")
        
        return False

    def is_matching_tab(self, title: str) -> bool:
        """Check if the window title matches the domain using Chrome-like pattern matching."""
        if not self.domain:  # For no-argument case
            return True
            
        if not title:  # Skip empty titles
            return False
            
        title_lower = title.lower()
        self.log(f"Analyzing title: {title_lower}")
        
        # Special case patterns for specific domains
        domain_specific_patterns = {
            "perplexity.ai": [
                "i'm looking for",  # Common in Perplexity search URLs
                "perplexity search",
                "perplexity - ",
                "perplexity ai",  # More generic match
                "perplexity.ai/search",  # Direct search URL match
            ]
        }
        
        # Base patterns for this domain
        base_patterns = [
            self.domain,  # Basic domain match without slash
            f"www.{self.domain}",  # www prefix without slash
            f"{self.domain}/",  # With slash
            f"www.{self.domain}/",  # www with slash
            f"https://{self.domain}",  # Full https without slash
            f"https://www.{self.domain}",  # Full https with www without slash
            self.domain.split('.')[0],  # First part of domain
        ]
        
        # Add domain-specific patterns if they exist
        if self.domain in domain_specific_patterns:
            base_patterns.extend(domain_specific_patterns[self.domain])
        
        # Get custom patterns from config and combine with base patterns
        url_patterns = self.config.get("url_patterns", {}).get(self.domain, [])
        title_patterns = self.config.get("title_patterns", {}).get(self.domain, [])
        all_patterns = base_patterns + url_patterns + title_patterns
        
        self.log(f"Checking against patterns: {all_patterns}")
        
        # Check all patterns
        for pattern in all_patterns:
            pattern = pattern.lower()
            if pattern in title_lower:
                self.log(f"Matched pattern: {pattern}")
                return True
        
        self.log(f"No patterns matched")
        return False

        # Add documentation for return value
    def focus_tab(self) -> bool:
        """Main function to find and focus the target tab.
        
        Returns:
            bool: True if successful (either focused existing tab or opened new one),
                  False if operation failed
        """
        """Main function to find and focus the target tab."""
        start_time = time.time()
        
        # Find existing Chrome window
        hwnd = self.find_chrome_window()
        if hwnd:
            # Focus the window
            if self.focus_window(hwnd):
                # Check if we're already on the target domain
                current_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                already_focused = self.is_matching_tab(current_title)
                
                if already_focused:
                    # If already focused, open a new tab with the domain
                    self.log("Already focused on target domain, opening new tab")
                    launch_result = self.launch_chrome(new_tab=True)
                    if launch_result:
                        self.log("Successfully opened new tab")
                    else:
                        self.log("Failed to open new tab")
                    return launch_result
                
                # Not focused on target domain, try to find target tab
                if self.cycle_through_tabs():
                    self.log(f"Found and focused tab in {time.time() - start_time:.2f}s")
                    return True
                else:
                    # Tab not found, open new one
                    self.log("Tab not found, opening in new tab")
                    launch_result = self.launch_chrome(new_tab=True)
                    if launch_result:
                        self.log("Successfully opened new tab")
                    else:
                        self.log("Failed to open new tab")
                    return launch_result
        
        # No Chrome window found, launch new instance
        self.log("No Chrome window found, launching new instance")
        launch_result = self.launch_chrome()
        if launch_result:
            self.log("Successfully launched new Chrome instance")
        else:
            self.log("Failed to launch new Chrome instance")
        return launch_result

def main():
    # If no arguments, just switch to Chrome
    if len(sys.argv) == 1:
        switcher = ChromeTabSwitcher("")  # Empty domain for Chrome-only focus
        try:
            # Just find and focus Chrome
            hwnd = switcher.find_chrome_window()
            if hwnd:
                success = switcher.focus_window(hwnd)
            else:
                # Launch new Chrome window if none exists
                success = switcher.launch_chrome()
            sys.exit(0 if success else 1)
        except Exception as e:
            switcher.log(f"Critical error: {e}")
            sys.exit(1)
    
    domain = sys.argv[1]
    switcher = ChromeTabSwitcher(domain)
    try:
        success = switcher.focus_tab()
        sys.exit(0 if success else 1)
    except Exception as e:
        switcher.log(f"Critical error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()