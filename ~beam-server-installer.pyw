import os
import requests
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import time # Import the time module for delays

# Conditional import for pywin32 (Windows-specific for shortcuts)
try:
    import win32com.client
    WINDOWS_SHORTCUT_SUPPORT = True
except ImportError:
    WINDOWS_SHORTCUT_SUPPORT = False
    print("[WARNING] 'pywin32' not found. Shortcut creation will be skipped on non-Windows systems or if not installed.")


# --- CONFIGURABLE VARIABLES ---
# The name of the main installation folder
INSTALL_FOLDER_NAME = "NameMe"
# URL to download the BeamMP Server executable
SERVER_DOWNLOAD_URL = "https://github.com/BeamMP/BeamMP-Server/releases/latest/download/BeamMP-Server.exe"
# The full path to the source folder containing BeamNG.drive mods (ZIP files)
# IMPORTANT: This path is Windows-specific. Adjust if on another OS or your path differs.
MOD_SOURCE_FOLDER = r"D:\Files\Important\AppData\Games\Beamng.drive\0.36\mods"
# Base path where the INSTALL_FOLDER_NAME will be created.
# By default, it's the directory where this script is run.
# Change this if you want to install it elsewhere (e.g., os.path.expanduser("~/Desktop"))
BASE_INSTALL_PATH = os.getcwd() # Installer will create NameMe inside this base path
# Name for the shortcut file
SHORTCUT_NAME = f"{INSTALL_FOLDER_NAME} - Shortcut.lnk" # .lnk is for Windows
# List of files to check for and move if found in the script's directory
FILES_TO_CHECK_AND_MOVE = ["BeamMP-Server.exe", "ServerConfig.toml"]
# --- END CONFIGURABLE VARIABLES ---


class BeamMPInstallerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BeamMP Server Installer")
        self.geometry("600x550") # Increased height for progress bar
        self.resizable(False, False)

        # Apply a modern theme
        self.style = ttk.Style(self)
        self.style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'

        self.full_install_path = os.path.join(BASE_INSTALL_PATH, INSTALL_FOLDER_NAME)
        self.client_mods_path = None # Will be set after directory creation

        self.create_widgets()

    def create_widgets(self):
        # Frame for controls
        control_frame = ttk.Frame(self, padding="15 15 15 15")
        control_frame.pack(fill=tk.BOTH, expand=False)

        # Installation Path Display
        ttk.Label(control_frame, text="Installation Directory:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.install_path_label = ttk.Label(control_frame, text=self.full_install_path, wraplength=400)
        self.install_path_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        # Modded Server Checkbox
        self.is_modded_var = tk.BooleanVar(value=False)
        self.modded_checkbox = ttk.Checkbutton(control_frame, text="Is this a modded server?", variable=self.is_modded_var)
        self.modded_checkbox.grid(row=2, column=0, sticky="w", pady=(0, 15))

        # Install Button
        self.install_button = ttk.Button(control_frame, text="Start Installation", command=self.start_installation)
        self.install_button.grid(row=3, column=0, columnspan=2, pady=(10, 0))

        # Progress Bar
        self.progress_bar = ttk.Progressbar(control_frame, orient='horizontal', length=500, mode='determinate')
        self.progress_bar.grid(row=4, column=0, columnspan=2, pady=(15, 10))
        self.progress_label = ttk.Label(control_frame, text="Ready to start...")
        self.progress_label.grid(row=5, column=0, columnspan=2, sticky="ew")

        # Separator
        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=10)

        # Log output area
        log_frame = ttk.LabelFrame(self, text="Installation Log", padding="10 10 10 10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15, state='disabled', font=("Monospace", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_config('info', foreground='blue')
        self.log_text.tag_config('error', foreground='red')
        self.log_text.tag_config('success', foreground='green')
        self.log_text.tag_config('warning', foreground='orange')

    def log_message(self, message, message_type='info'):
        """Inserts a message into the scrolled text widget."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n", message_type)
        self.log_text.see(tk.END) # Auto-scroll to the end
        self.log_text.config(state='disabled')
        self.update_idletasks() # Update GUI immediately

    def update_progress(self, value, status_text=""):
        """Updates the progress bar and status label."""
        self.progress_bar['value'] = value
        self.progress_label.config(text=status_text)
        self.update_idletasks()

    def start_installation(self):
        # Disable controls and clear log
        self.install_button.config(state='disabled', text="Installing...")
        self.modded_checkbox.config(state='disabled')
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END) # Clear previous log
        self.log_text.config(state='disabled')
        self.update_progress(0, "Starting installation...")
        self.log_message("Starting BeamMP Server Installer (GUI Version)", 'info')

        # 1. Create main installation folder and subfolders
        self.log_message("Creating installation directories...", 'info')
        self.update_progress(5, "Creating directories...")
        self.full_install_path, self.client_mods_path = self._create_directories(BASE_INSTALL_PATH, INSTALL_FOLDER_NAME)
        if not self.full_install_path:
            self.log_message("Installation aborted due to directory creation error.", 'error')
            self._installation_complete(False)
            return

        # 2. Download BeamMP-Server.exe
        server_exe_name = os.path.basename(SERVER_DOWNLOAD_URL)
        server_exe_destination = os.path.join(self.full_install_path, server_exe_name)
        self.log_message(f"Downloading {server_exe_name}...", 'info')
        self.update_progress(10, "Downloading server executable...")
        if not self._download_file(SERVER_DOWNLOAD_URL, server_exe_destination):
            self.log_message("Installation aborted due to download error.", 'error')
            self._installation_complete(False)
            return
        self.update_progress(50, "Download complete.")

        # 3. Open the downloaded .exe file
        self.log_message(f"Launching {server_exe_name}...", 'info')
        self.update_progress(55, "Launching server executable...")
        if not self._execute_exe(server_exe_destination):
            self.log_message("Could not launch the server executable, but proceeding with other steps if possible.", 'warning')

        # 4. Copy mods if selected
        if self.is_modded_var.get():
            self.log_message("Modded server option selected. Copying mods...", 'info')
            self.update_progress(60, "Copying mods...")
            self._copy_zip_files(MOD_SOURCE_FOLDER, self.client_mods_path)
            self.update_progress(85, "Mods copied.")
        else:
            self.log_message("Modded server option not selected. Skipping mod copying.", 'info')
            self.update_progress(85, "Skipping mod copying.") # Still update progress

        # 5. Create shortcut
        self.log_message("Creating shortcut...", 'info')
        self.update_progress(90, "Creating shortcut...")
        self._create_shortcut(server_exe_destination, BASE_INSTALL_PATH, SHORTCUT_NAME) # Shortcut in BASE_INSTALL_PATH

        # 6. Delay and move existing files (LAST STEP)
        self.log_message("Waiting 1 second before moving existing files...", 'info')
        self.update_progress(92, "Preparing to move existing files...")
        time.sleep(1) # Wait for 1 second

        self.log_message("Moving existing server files...", 'info')
        # Explicitly get the directory of the currently running script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self._check_and_move_existing_files(script_dir, self.full_install_path)
        self.update_progress(98, "Existing files moved.")

        self._installation_complete(True)

    def _create_directories(self, base_path, folder_name):
        """Creates the main installation directory and subdirectories."""
        full_install_path = os.path.join(base_path, folder_name)
        resources_path = os.path.join(full_install_path, "Resources")
        client_path = os.path.join(resources_path, "Client")

        try:
            os.makedirs(client_path, exist_ok=True)
            self.log_message(f"Created directories: {full_install_path}, {resources_path}, {client_path}", 'success')
            return full_install_path, client_path
        except OSError as e:
            self.log_message(f"Error creating directories: {e}", 'error')
            return None, None

    def _check_and_move_existing_files(self, source_dir, destination_dir):
        """Checks for specific files in source_dir and moves them to destination_dir."""
        moved_any = False
        self.log_message(f"Checking for existing files in '{source_dir}'...", 'info')
        for filename in FILES_TO_CHECK_AND_MOVE:
            source_path = os.path.join(source_dir, filename)
            destination_path = os.path.join(destination_dir, filename)
            if os.path.exists(source_path):
                try:
                    shutil.move(source_path, destination_path)
                    self.log_message(f"Moved existing file '{filename}' to '{destination_dir}'", 'success')
                    moved_any = True
                except Exception as e:
                    self.log_message(f"Error moving existing file '{filename}': {e}", 'error')
            else:
                self.log_message(f"File '{filename}' not found in '{source_dir}'. Skipping move.", 'info')

        if not moved_any:
            self.log_message("No specified existing server files were moved.", 'info')


    def _download_file(self, url, destination_path):
        """Downloads a file from a URL to a specified destination."""
        try:
            self.log_message(f"Downloading from: {url}")
            response = requests.get(url, stream=True)
            response.raise_for_status() # Raise an exception for HTTP errors

            total_size = int(response.headers.get('content-length', 0))
            bytes_downloaded = 0
            
            with open(destination_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
                    bytes_downloaded += len(chunk)
                    # Update download progress relative to this step's progress (10% to 50%)
                    if total_size > 0:
                        download_progress_percentage = (bytes_downloaded / total_size) * (50 - 10)
                        self.update_progress(10 + download_progress_percentage, f"Downloading: {bytes_downloaded}/{total_size} bytes")

            self.log_message(f"Downloaded successfully to: {destination_path}", 'success')
            return True
        except requests.exceptions.RequestException as e:
            self.log_message(f"Error downloading file: {e}", 'error')
            return False
        except IOError as e:
            self.log_message(f"Error saving downloaded file: {e}", 'error')
            return False

    def _execute_exe(self, exe_path):
        """Executes an .exe file without waiting for it to finish."""
        try:
            self.log_message(f"Opening executable: {exe_path}")
            if sys.platform.startswith('win'):
                subprocess.Popen([exe_path], shell=True)
            else:
                subprocess.Popen([exe_path])
            self.log_message("Executable launched. Continuing with installation...", 'info')
            return True
        except OSError as e:
            self.log_message(f"Error executing file '{exe_path}': {e}", 'error')
            self.log_message("Please ensure the file is executable and you have permissions.", 'error')
            return False

    def _copy_zip_files(self, source_folder, destination_folder):
        """Copies all .zip files from a source to a destination folder."""
        if not os.path.isdir(source_folder):
            self.log_message(f"Mod source folder not found: {source_folder}", 'error')
            return False

        copied_count = 0
        self.log_message(f"Searching for .zip files in: {source_folder}")
        
        zip_files = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f)) and f.lower().endswith(".zip")]
        total_zip_files = len(zip_files)

        for i, item_name in enumerate(zip_files):
            source_item_path = os.path.join(source_folder, item_name)
            try:
                shutil.copy2(source_item_path, destination_folder)
                self.log_message(f"Copied: {item_name}")
                copied_count += 1
                # Update progress for mod copying (60% to 85%)
                copy_progress_percentage = (i + 1) / total_zip_files * (85 - 60) if total_zip_files > 0 else 0
                self.update_progress(60 + copy_progress_percentage, f"Copying mods: {copied_count}/{total_zip_files}")
            except Exception as e:
                self.log_message(f"Error copying {item_name}: {e}", 'error')
        
        if copied_count > 0:
            self.log_message(f"Successfully copied {copied_count} mod(s) to: {destination_folder}", 'success')
        else:
            self.log_message("No .zip files found to copy or an error occurred during copying.", 'info')
        return True

    def _create_shortcut(self, target_path, shortcut_dir, shortcut_name):
        """Creates a shortcut to the target_path in the shortcut_dir."""
        if sys.platform.startswith('win') and WINDOWS_SHORTCUT_SUPPORT:
            shortcut_path = os.path.join(shortcut_dir, shortcut_name)
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target_path
                shortcut.WorkingDirectory = os.path.dirname(target_path) # Set working directory to EXE's folder
                shortcut.IconLocation = target_path # Use the EXE's icon
                shortcut.Save()
                self.log_message(f"Created shortcut at: {shortcut_path}", 'success')
                return True
            except Exception as e:
                self.log_message(f"Error creating shortcut: {e}", 'error')
                self.log_message("Ensure 'pywin32' is installed and you have permissions.", 'error')
                return False
        else:
            self.log_message("Shortcut creation is only supported on Windows with 'pywin32' installed. Skipping.", 'warning')
            return False

    def _installation_complete(self, success):
        """Finalizes the installation process."""
        self.install_button.config(state='normal', text="Start Installation")
        self.modded_checkbox.config(state='normal')
        self.update_progress(100, "Installation Complete!")
        
        if success:
            self.log_message("\nInstallation process completed successfully!", 'success')
            messagebox.showinfo("Installation Complete", "BeamMP Server installation finished successfully!")
        else:
            self.log_message("\nInstallation process encountered errors.", 'error')
            messagebox.showerror("Installation Error", "BeamMP Server installation encountered errors. Please check the log.")


if __name__ == "__main__":
    app = BeamMPInstallerGUI()
    app.mainloop()
