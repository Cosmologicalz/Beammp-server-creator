import os
import requests
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import time # Import the time module for delays
import webbrowser
import random
import string


# Conditional import for pywin32 (Windows-specific for shortcuts)
try:
    import win32com.client
    WINDOWS_SHORTCUT_SUPPORT = True
except ImportError:
    WINDOWS_SHORTCUT_SUPPORT = False
    # Suppress print warning to avoid console output for .pyw file,
    # as the GUI handles the user notification.

def generate_random_code():
    """
    Generates a random 10-character code consisting of digits and letters.
    """
    characters = string.digits + string.ascii_letters
    random_code = ''.join(random.choices(characters, k=10))
    return random_code

# Generate the random code once to be used for both shortcut and text file
GLOBAL_SHORTCUT_CODE = generate_random_code()

# --- CONFIGURABLE VARIABLES ---
# Current version of the installer
VERSION = "v0.2.3" # Incrementing version for size display feature
# The name of the main installation folder
INSTALL_FOLDER_NAME = "NameMe"
# URL to download the BeamMP Server executable
SERVER_DOWNLOAD_URL = "https://github.com/BeamMP/BeamMP-Server/releases/latest/download/BeamMP-Server.exe"
# The full path to the source folder containing BeamNG.drive mods (ZIP files)
# IMPORTANT: This path is Windows-specific. Adjust if on another OS or your path differs.
MOD_SOURCE_FOLDER = r"D:\Files\Important\AppData\Games\Beamng.drive\0.36\mods\repo"
DEBUG = False
# Base path where the INSTALL_FOLDER_NAME will be created.
# By default, it's the directory where this script is run.
# Change this if you want to install it elsewhere (e.g., os.path.expanduser("~/Desktop"))
BASE_INSTALL_PATH = os.getcwd() # Installer will create NameMe inside this base path
# Name for the shortcut file, now using the globally generated code
SHORTCUT_NAME = f"{INSTALL_FOLDER_NAME} - DO NOT DELETE CODE [{GLOBAL_SHORTCUT_CODE}] .lnk" # .lnk is for Windows
# List of files to check for and move if found in the script's directory
FILES_TO_CHECK_AND_MOVE = ["BeamMP-Server.exe", "ServerConfig.toml"]
# --- END CONFIGURABLE VARIABLES ---

if DEBUG:
    MOD_SOURCE_FOLDER = r"debugfolder"

def get_human_readable_size(size_bytes):
    """Converts a size in bytes to a human-readable format (B, KB, MB, GB)."""
    if size_bytes < 1024:
        return f"{size_bytes} Bytes"
    elif size_bytes < 1024**2:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024**3:
        return f"{size_bytes / (1024**2):.2f} MB"
    else:
        return f"{size_bytes / (1024**3):.2f} GB"

class BeamMPInstallerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"BeamMP Server Installer {VERSION}") # Display version in title
        self.geometry("600x550") # Increased height for progress bar
        self.resizable(False, False)

        # Apply a modern theme
        self.style = ttk.Style(self)
        self.style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'

        self.full_install_path = os.path.join(BASE_INSTALL_PATH, INSTALL_FOLDER_NAME)
        self.client_mods_path = None # Will be set after directory creation

        # Initialize mod folder existence flag and error message
        self.mod_folder_exists = False
        self.initial_mod_folder_error = None # To store error message if check fails early

        # Check if the mod source folder exists and is accessible on startup
        try:
            if os.path.isdir(MOD_SOURCE_FOLDER):
                self.mod_folder_exists = True
            else:
                # Path exists but is not a directory, or path does not exist
                self.initial_mod_folder_error = f"Mod folder does not exist or is not a directory: '{MOD_SOURCE_FOLDER}'"
        except Exception as e: # Catch any other exceptions (e.g., permission denied)
            self.initial_mod_folder_error = f"Error accessing mod folder '{MOD_SOURCE_FOLDER}': {e}"

        self.create_widgets()

        # Display initial error if any occurred during mod folder check
        if self.initial_mod_folder_error:
            messagebox.showerror(
                "Startup Warning: Mod Folder Issue",
                self.initial_mod_folder_error + "\n\nMod copying functionality will be disabled. "
                "Please ensure the path is correct and the installer has permissions to access it."
            )
            # Also log this to the scrolled text
            self.log_message(f"Warning: {self.initial_mod_folder_error}. Mod copying disabled.", 'error')


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

        # Disable checkbox if mod folder doesn't exist or had an initial error
        if not self.mod_folder_exists:
            self.modded_checkbox.config(state='disabled')
            # Bind a click event to show a warning when disabled
            self.modded_checkbox.bind("<Button-1>", self._show_mod_folder_warning)
            # Initial warning message already handled in __init__ if there was an error.
            # If no error but just not a dir, log it here.
            if not self.initial_mod_folder_error: 
                self.log_message(f"Warning: Mod folder not found at '{MOD_SOURCE_FOLDER}'. Mod copying will be disabled.", 'warning')

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

    def _show_mod_folder_warning(self, event=None):
        """Shows a warning message if the mod folder doesn't exist or couldn't be accessed."""
        error_detail_message = ""
        if self.initial_mod_folder_error:
            error_detail_message = f"\n\nDetails: {self.initial_mod_folder_error}"

        messagebox.showwarning(
            "Mod Folder Not Found / Accessible",
            f"The specified mod folder does not exist or could not be accessed:\n'{MOD_SOURCE_FOLDER}'{error_detail_message}\n\n"
            "Please ensure BeamNG.drive is installed, the folder path is correct, "
            "and that the installer has necessary permissions to read this directory. "
            "Then, restart the installer."
        )
        # Prevent the checkbox state from changing if it was disabled
        return "break" # Prevents default event handling for Tkinter widgets

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

        # NEW: Create bsi_code.txt file
        self.log_message("Creating bsi_code.txt file...", 'info')
        self.update_progress(7, "Creating code file...")
        self._create_bsi_code_file(self.full_install_path, GLOBAL_SHORTCUT_CODE)
        self.update_progress(9, "Code file created.")

        # 2. Download BeamMP-Server.exe
        server_exe_name = os.path.basename(SERVER_DOWNLOAD_URL)
        server_exe_destination = os.path.join(self.full_install_path, server_exe_name)
        self.log_message(f"Downloading {server_exe_name}...", 'info')
        self.update_progress(10, "Downloading server executable...")
        if not self._download_file(SERVER_DOWNLOAD_URL, server_exe_destination):
            self.log_message("Installation aborted due to download error.", 'error')
            self._installation_complete(False)
            return
        self.update_progress(45, "Download complete.") # Adjusted progress

        # 3. Open the downloaded .exe file
        self.log_message(f"Launching {server_exe_name}...", 'info')
        self.update_progress(50, "Launching server executable...") # Adjusted progress
        if not self._execute_exe(server_exe_destination):
            self.log_message("Could not launch the server executable, but proceeding with other steps if possible.", 'warning')

        # 4. Copy mods if selected and mod folder exists
        # Ensure that self.mod_folder_exists is true before attempting to copy
        if self.is_modded_var.get() and self.mod_folder_exists:
            self.log_message("Modded server option selected. Copying files from mod folder...", 'info')
            self.update_progress(55, "Copying mod files...") # Adjusted progress
            # Changed to copy all files, not just .zip
            total_copied_size = self._copy_all_files_from_folder(MOD_SOURCE_FOLDER, self.client_mods_path) 
            self.log_message(f"Data size of mods successfully installed: {get_human_readable_size(total_copied_size)}", 'success') # NEW: Display size
            self.update_progress(75, "Mod files copied.") # Adjusted progress
        else:
            self.log_message("Modded server option not selected or mod folder missing/inaccessible. Skipping mod copying.", 'info')
            self.update_progress(75, "Skipping mod copying.") # Still update progress

        # 5. Create shortcut
        self.log_message("Creating shortcut...", 'info')
        self.update_progress(80, "Creating shortcut...") # Adjusted progress
        # Shortcut will be created in BASE_INSTALL_PATH, not the NameMe folder
        shortcut_path = os.path.join(BASE_INSTALL_PATH, SHORTCUT_NAME)
        self._create_shortcut(server_exe_destination, BASE_INSTALL_PATH, SHORTCUT_NAME) 

        # 6. Delay and move existing files (LAST STEP before verification)
        self.log_message("Waiting 1 second before moving existing files...", 'info')
        self.update_progress(85, "Preparing to move existing files...") # Adjusted progress
        time.sleep(1) # Wait for 1 second

        self.log_message("Moving existing server files...", 'info')
        # Explicitly get the directory of the currently running script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self._check_and_move_existing_files(script_dir, self.full_install_path)
        self.update_progress(90, "Existing files moved.") # Adjusted progress

        # 7. Post-installation verification (NEW STEP)
        self.log_message("Verifying installed files...", 'info')
        self.update_progress(95, "Verifying installation...")
        verification_success = self._verify_installation_files(
            self.full_install_path, 
            server_exe_destination, 
            shortcut_path
        )
        self.update_progress(99, "Verification complete.")

        self._installation_complete(verification_success)

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
            
    def _create_bsi_code_file(self, install_path, code):
        """Creates a bsi_code.txt file with the generated code."""
        file_path = os.path.join(install_path, "bsi_code.txt")
        try:
            with open(file_path, 'w') as f:
                f.write(f"generated_bsi_shortcut_code = {code}\n")
            self.log_message(f"Created '{os.path.basename(file_path)}' successfully at '{install_path}'", 'success')
            return True
        except IOError as e:
            self.log_message(f"Error creating bsi_code.txt: {e}", 'error')
            return False

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
                    # Update download progress relative to this step's progress (10% to 45%)
                    if total_size > 0:
                        download_progress_percentage = (bytes_downloaded / total_size) * (45 - 10)
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

    def _copy_all_files_from_folder(self, source_folder, destination_folder):
        """Copies all files from a source to a destination folder and returns their total size."""
        # This check is technically redundant due to self.mod_folder_exists check earlier,
        # but kept for robustness within the function itself.
        if not os.path.isdir(source_folder):
            self.log_message(f"Source folder not found for copying files: {source_folder}", 'error')
            return 0 # Return 0 size if source folder doesn't exist or is not a directory

        copied_count = 0
        total_copied_size = 0 # Initialize total size
        self.log_message(f"Searching for files in: {source_folder}")
        
        try:
            # Get a list of all files (not directories) in the source folder
            files_to_copy = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f))]
        except Exception as e:
            self.log_message(f"Error listing files in source folder '{source_folder}': {e}", 'error')
            self.log_message("File copying failed due to folder access issue.", 'error')
            return 0 # Return 0 size on error

        total_files = len(files_to_copy)

        for i, item_name in enumerate(files_to_copy):
            source_item_path = os.path.join(source_folder, item_name)
            try:
                shutil.copy2(source_item_path, destination_folder)
                file_size = os.path.getsize(source_item_path) # Get size of copied file
                total_copied_size += file_size # Add to total
                self.log_message(f"Copied: {item_name} ({get_human_readable_size(file_size)})", 'success') # Added size to log
                copied_count += 1
                # Update progress for mod copying (55% to 75%)
                copy_progress_percentage = (i + 1) / total_files * (75 - 55) if total_files > 0 else 0
                self.update_progress(55 + copy_progress_percentage, f"Copying files: {copied_count}/{total_files}")
            except Exception as e:
                self.log_message(f"Error copying {item_name}: {e}", 'error')
        
        if copied_count > 0:
            self.log_message(f"Successfully copied {copied_count} file(s) to: {destination_folder}", 'success')
        else:
            self.log_message("No files found to copy or an error occurred during copying.", 'info')
        
        return total_copied_size # Return the total size

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

    def _verify_installation_files(self, full_install_path, server_exe_destination, shortcut_path):
        """Verifies the presence of key installed files and folders."""
        overall_success = True
        self.log_message("--- Verifying Installation ---", 'info')

        # Check main folder
        if os.path.isdir(full_install_path):
            self.log_message(f"'{INSTALL_FOLDER_NAME}' folder exists.", 'success')
        else:
            self.log_message(f"Error: '{INSTALL_FOLDER_NAME}' folder NOT found.", 'error')
            overall_success = False

        # Check Resources folder
        resources_path = os.path.join(full_install_path, "Resources")
        if os.path.isdir(resources_path):
            self.log_message(f"'{os.path.basename(resources_path)}' folder exists.", 'success')
        else:
            self.log_message(f"Error: '{os.path.basename(resources_path)}' folder NOT found.", 'error')
            overall_success = False

        # Check Client folder
        client_path = os.path.join(resources_path, "Client")
        if os.path.isdir(client_path):
            self.log_message(f"'{os.path.basename(client_path)}' folder exists.", 'success')
        else:
            self.log_message(f"Error: '{os.path.basename(client_path)}' folder NOT found.", 'error')
            overall_success = False

        # Check BeamMP-Server.exe
        if os.path.exists(server_exe_destination):
            self.log_message(f"'{os.path.basename(server_exe_destination)}' executable exists.", 'success')
        else:
            self.log_message(f"Error: '{os.path.basename(server_exe_destination)}' executable NOT found.", 'error')
            overall_success = False

        # Check ServerConfig.toml (if it was one of the files to check/move)
        server_config_name = "ServerConfig.toml"
        if server_config_name in FILES_TO_CHECK_AND_MOVE:
            toml_path = os.path.join(full_install_path, server_config_name)
            if os.path.exists(toml_path):
                self.log_message(f"'{server_config_name}' file exists.", 'success')
            else:
                self.log_message(f"Warning: '{server_config_name}' file NOT found (might not have been present in source or moved).", 'warning')
        else:
            self.log_message(f"Note: '{server_config_name}' was not configured for specific verification.", 'info')

        # Check shortcut
        if os.path.exists(shortcut_path):
            self.log_message(f"Shortcut '{os.path.basename(shortcut_path)}' exists.", 'success')
        else:
            self.log_message(f"Error: Shortcut '{os.path.basename(shortcut_path)}' NOT found.", 'error')
            overall_success = False
        
        # Check bsi_code.txt file (NEW VERIFICATION)
        bsi_code_file_path = os.path.join(full_install_path, "bsi_code.txt")
        if os.path.exists(bsi_code_file_path):
            self.log_message(f"'bsi_code.txt' file exists.", 'success')
            try:
                with open(bsi_code_file_path, 'r') as f:
                    content = f.read().strip()
                expected_content = f"generated_bsi_shortcut_code = {GLOBAL_SHORTCUT_CODE}"
                if content == expected_content:
                    self.log_message(f"'bsi_code.txt' content matches expected code.", 'success')
                else:
                    self.log_message(f"Warning: 'bsi_code.txt' content mismatch. Expected: '{expected_content}', Found: '{content}'.", 'warning')
            except Exception as e:
                self.log_message(f"Error reading 'bsi_code.txt': {e}", 'error')
        else:
            self.log_message(f"Error: 'bsi_code.txt' file NOT found.", 'error')
            overall_success = False
            
        self.log_message("--- Verification Complete ---", 'info')
        return overall_success

    def _installation_complete(self, success):
        """Finalizes the installation process and handles auto-closing."""
        self.install_button.config(state='normal', text="Start Installation")
        self.modded_checkbox.config(state='normal')
        self.update_progress(100, "Installation Complete!")
        
        if success:
            self.log_message("\nInstallation process completed successfully!", 'success')
            messagebox.showinfo("Installation Complete", "BeamMP Server installation finished successfully!")
        else:
            self.log_message("\nInstallation process encountered errors. Please check the log for details.", 'error')
            messagebox.showerror("Installation Error", "BeamMP Server installation encountered errors. Please check the log.")

        # Auto-close after 1 second
        self.log_message("Installer will close in 1 second...", 'info')
        self.update_idletasks() # Ensure message appears before delay
        time.sleep(1)
        self.destroy() # Close the Tkinter window


if __name__ == "__main__":
    app = BeamMPInstallerGUI()
    app.mainloop()
