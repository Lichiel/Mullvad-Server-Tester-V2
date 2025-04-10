# automated_tester.py
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import platform
import threading
import logging
import logging.handlers
import subprocess
import json # For parsing speedtest output

# --- Setup Logging ---
try:
    # Use a different log file name for this application
    log_file_path = os.path.expanduser("~/mullvad_tester.log")

    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    # Use rotating file handler to limit log file size
    log_handler = logging.handlers.RotatingFileHandler(
        log_file_path, maxBytes=2*1024*1024, backupCount=1, encoding='utf-8'
    )
    log_handler.setFormatter(logging.Formatter(log_format))

    logging.basicConfig(
        level=logging.INFO, # Default level, can be changed
        # level=logging.DEBUG, # Uncomment for more detailed logs
        format=log_format,
        handlers=[
            log_handler,
            logging.StreamHandler(sys.stdout) # Also log to console
        ]
    )
    logger = logging.getLogger(__name__)
    logger.info(f"Logging initialized. Log file: {log_file_path}")

except Exception as e:
     # Basic fallback logging if setup fails
     print(f"FATAL: Failed to initialize logging: {e}", file=sys.stderr)
     logging.basicConfig(level=logging.ERROR)
     logger = logging.getLogger(__name__)
     logger.error(f"Logging setup failed: {e}")


# --- Add project root and example dir to path ---
try:
    project_root = os.path.dirname(os.path.abspath(__file__))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)
    logger.debug(f"Project root added to sys.path: {project_root}")

    # --- Import GUI Application (from new file) ---
    # We will create tester_gui.py next
    from tester_gui import TesterApp
    # Import config functions explicitly
    from config import get_log_path
except ImportError as e:
     logger.exception("Failed to import TesterApp or config. Critical dependency missing? Ensure tester_gui.py and config.py exist.")
     # Use basic tkinter messagebox if available
     try:
         root = tk.Tk()
         root.withdraw()
         messagebox.showerror("Startup Error", f"Failed to load application components:\n{e}\n\nPlease ensure tester_gui.py and config.py are present.")
     except Exception:
         print(f"FATAL STARTUP ERROR: Failed to load application components: {e}", file=sys.stderr)
     sys.exit(1)
except Exception as e:
     logger.exception("An unexpected error occurred during initial imports.")
     try:
         root = tk.Tk()
         root.withdraw()
         messagebox.showerror("Startup Error", f"An unexpected error occurred on startup:\n{e}")
     except Exception:
          print(f"FATAL STARTUP ERROR: An unexpected error occurred on startup: {e}", file=sys.stderr)
     sys.exit(1)


# --- Platform Specific Setup (reuse from example/main.py) ---
def set_dpi_awareness():
    """Set DPI awareness, primarily for Windows."""
    if platform.system() == 'Windows':
        try:
            import ctypes
            awareness = ctypes.c_int()
            errorCode = ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
            logger.info(f"Initial DPI Awareness: {awareness.value}")
            # Set Per-Monitor DPI Awareness v2 if available (Win 10 1703+)
            # PROCESS_PER_MONITOR_DPI_AWARE = 2
            if awareness.value < 2: # If not already Per-Monitor aware
                errorCode = ctypes.windll.shcore.SetProcessDpiAwareness(2)
                if errorCode == 0: # S_OK
                    logger.info("Successfully set Per-Monitor DPI Awareness.")
                else:
                    logger.error(f"Failed to set Per-Monitor DPI Awareness, Error Code: {errorCode}")
        except ImportError: logger.warning("Could not import ctypes, cannot set DPI awareness.")
        except AttributeError: logger.warning("shcore.SetProcessDpiAwareness not found (requires Windows 8.1+).")
        except Exception as e: logger.exception(f"Error setting DPI awareness: {e}")


# --- Dependency Checks ---

def check_mullvad_cli() -> bool:
    """Check if the Mullvad CLI is installed and accessible."""
    logger.info("Checking for Mullvad CLI dependency...")
    try:
        result = subprocess.run(
            ['mullvad', 'version'],
            capture_output=True, text=True, check=False, timeout=5, encoding='utf-8', errors='ignore'
        )
        if result.returncode == 0:
            logger.info(f"Mullvad CLI check successful. Output: '{result.stdout.strip()}'")
            return True
        else:
            logger.error(f"Mullvad CLI check command failed. Return code: {result.returncode}, Stderr: '{result.stderr.strip()}'")
            return False
    except FileNotFoundError:
        logger.error("Mullvad CLI command ('mullvad') not found in PATH.")
        return False
    except subprocess.TimeoutExpired:
        logger.error("Mullvad CLI check command timed out.")
        return False
    except Exception as e:
        logger.exception(f"Unexpected error during Mullvad CLI check: {e}")
        return False

def check_speedtest_cli() -> bool:
    """Check if the Ookla Speedtest CLI is installed and accessible."""
    logger.info("Checking for Speedtest CLI dependency...")
    try:
        # Try running 'speedtest --version' first
        result = subprocess.run(
            ['speedtest', '--version'],
            capture_output=True, text=True, check=False, timeout=5, encoding='utf-8', errors='ignore'
        )
        output = result.stdout.strip() + result.stderr.strip()
        # Check for common version strings or Ookla name
        if 'Ookla' in output or 'Speedtest by Ookla' in output or 'speedtest-cli' in output.lower():
             logger.info(f"Speedtest CLI check successful via --version. Output: '{output}'")
             return True
        else:
             # If --version failed or didn't give expected output, try a minimal run
             logger.info("Speedtest --version check inconclusive, trying basic execution check...")
             # Use --accept-license and --accept-gdpr to bypass initial prompts
             result_basic = subprocess.run(
                 ['speedtest', '--format=json', '--accept-license', '--accept-gdpr'],
                 capture_output=True, text=True, check=False, timeout=15, encoding='utf-8', errors='ignore' # Slightly longer timeout
             )
             # Check if it produced JSON-like output or exited reasonably
             stdout_basic = result_basic.stdout.strip()
             stderr_basic = result_basic.stderr.strip()
             if result_basic.returncode == 0 and stdout_basic.startswith('{') and stdout_basic.endswith('}'):
                  logger.info(f"Speedtest CLI basic execution check successful (produced JSON).")
                  return True
             elif 'error' in stderr_basic.lower() and 'configuration' in stderr_basic.lower():
                  logger.warning(f"Speedtest CLI ran but reported a configuration error, assuming installed: {stderr_basic}")
                  return True # Assume installed, config issue is separate
             elif result_basic.returncode != 0 and ('license' in stderr_basic.lower() or 'terms' in stderr_basic.lower()):
                  logger.warning(f"Speedtest CLI ran but requires license acceptance interactively. Assuming installed.")
                  # Might need user to run `speedtest` once manually first
                  return True # Assume installed, needs manual first run
             else:
                  logger.error(f"Speedtest CLI basic execution check failed. Return code: {result_basic.returncode}, Stdout: '{stdout_basic[:100]}...', Stderr: '{stderr_basic[:100]}...'")
                  return False

    except FileNotFoundError:
        logger.error("Speedtest CLI command ('speedtest') not found in PATH.")
        return False
    except subprocess.TimeoutExpired:
        logger.error("Speedtest CLI check command timed out.")
        return False
    except Exception as e:
        logger.exception(f"Unexpected error during Speedtest CLI check: {e}")
        return False

# --- Main Execution ---

def main():
    """Main entry point for the Mullvad Automated Tester application."""
    logger.info("--- Mullvad Automated Tester Application Starting ---")
    set_dpi_awareness()

    # Check critical dependencies
    mullvad_ok = check_mullvad_cli()
    speedtest_ok = check_speedtest_cli()

    if not mullvad_ok:
        logger.critical("Mullvad CLI dependency check failed. Application cannot continue.")
        # Attempt to show Tkinter message box
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Dependency Error",
                "Mullvad CLI ('mullvad') not found or not working.\n\n"
                "Please ensure the Mullvad VPN client is installed correctly "
                "and the 'mullvad' command is accessible in your system's PATH.",
                parent=root
            )
        except Exception:
             print("FATAL ERROR: Mullvad CLI not found or not working.", file=sys.stderr)
        sys.exit(1)

    if not speedtest_ok:
        logger.critical("Speedtest CLI dependency check failed. Application cannot continue.")
        install_instructions = ("Please install the Ookla Speedtest CLI.\n\n"
                                "On macOS with Homebrew:\n"
                                "`brew tap teamookla/speedtest`\n"
                                "`brew install speedtest --force`\n\n"
                                "For other systems, visit speedtest.net/apps/cli\n\n"
                                "You may need to run `speedtest` once manually in your terminal "
                                "to accept the license agreement.")
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Dependency Error",
                f"Ookla Speedtest CLI ('speedtest') not found or not working.\n\n{install_instructions}",
                parent=root
            )
        except Exception:
             print(f"FATAL ERROR: Ookla Speedtest CLI not found or not working.\n{install_instructions}", file=sys.stderr)
        sys.exit(1)


    # --- Initialize Tkinter Root ---
    root = tk.Tk()
    root.withdraw() # Hide the window initially

    root.title("Mullvad Automated Tester")
    # Adjust initial size as needed
    root.minsize(900, 650)

    # --- Initialize and Run Application ---
    try:
        # Instantiate the main application class from tester_gui.py
        app = TesterApp(root)
        root.deiconify() # Show the window after initialization
        logger.info("Starting Tkinter main loop...")
        root.mainloop()
        logger.info("Tkinter main loop finished.")

    except Exception as e:
         # Try to get log path from config, fallback to initial path
         try:
             # Ensure get_log_path is imported correctly before calling
             from config import get_log_path
             log_path_final = get_log_path()
         except NameError: # If config import failed earlier
             log_path_final = log_file_path
         except Exception:
             log_path_final = log_file_path

         logger.exception("An unhandled exception occurred during application execution.")
         try:
             # Don't create a new root if one might exist
             messagebox.showerror("Fatal Error", f"An unexpected error occurred:\n{e}\n\nPlease check the log file:\n{log_path_final}", parent=root if root.winfo_exists() else None)
         except Exception:
              print(f"FATAL ERROR: An unexpected error occurred: {e}. Check log: {log_path_final}", file=sys.stderr)

         try:
             if root.winfo_exists():
                 root.destroy()
         except:
             pass # Ignore errors during destroy
         sys.exit(1)

    logger.info("--- Mullvad Automated Tester Application Exiting ---")


if __name__ == "__main__":
    # Add guard for multiprocessing on Windows if ever needed
    # if platform.system() == "Windows":
    #     import multiprocessing
    #     multiprocessing.freeze_support()
    main()
