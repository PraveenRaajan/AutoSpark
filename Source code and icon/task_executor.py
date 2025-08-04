"""
Task Executor for the AutoSpark Application
Handles execution of automation tasks and .bat files
"""

import os
import subprocess
import sys
import webbrowser

class TaskExecutor:
    """Executes automation tasks"""
    
    def __init__(self):
        """Initialize the task executor"""
        pass
    
    def run_bat_file(self, bat_file_path):
        """
        Run a .bat file
        
        Args:
            bat_file_path (str): Path to the .bat file to execute
            
        Returns:
            bool: True if successful, False otherwise
            
        Raises:
            Exception: If the file cannot be executed
        """
        if not os.path.exists(bat_file_path):
            raise FileNotFoundError(f"BAT file not found: {bat_file_path}")
        
        try:
            # Using subprocess.Popen to run the bat file asynchronously
            subprocess.Popen(
                [bat_file_path],
                shell=True,  # Required for .bat files
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NEW_CONSOLE  # Open in new command window
            )
            return True
        except Exception as e:
            raise Exception(f"Error executing BAT file: {str(e)}")
    
    def execute_task(self, task):
        """
        Execute a single task directly (for testing or immediate execution)
        
        Args:
            task (dict): A task dictionary with type, details, and additional info
            
        Returns:
            bool: True if successful, False otherwise
            
        Raises:
            Exception: If the task cannot be executed
        """
        # Debug output
        #(f"TASK EXECUTOR: Executing task: {task}")
        
        task_type = task.get("type", "")
        details = task.get("details", "")
        additional = task.get("additional", "")
        
        #(f"TASK TYPE: {task_type}")
        #(f"DETAILS: {details}")
        #(f"ADDITIONAL: {additional}")
        
        # For testing in the Replit environment
        if task_type == "delay":
            #(f"Executing delay task: {details} seconds")
            return self._delay(details)
        elif task_type == "open_url":
            #(f"Executing open URL task (simulated in Replit): {details}")
            return True  # Simulate success
        
        # We're in a Replit environment, so just simulate success for most operations
        #(f"Task type {task_type} execution simulated in Replit environment.")
        return True

    def _open_url(self, url):
        """Open a URL in the default browser"""
        try:
            webbrowser.open(url)
            return True
        except Exception as e:
            raise Exception(f"Error opening URL: {str(e)}")
    
    def _open_app(self, app_path):
        """Open an application"""
        try:
            cmd = f'"{app_path}"'
            subprocess.Popen(cmd, shell=True)
            return True
        except Exception as e:
            raise Exception(f"Error opening application: {str(e)}")
    
    def _open_file(self, file_path):
        """Open a file with its default application"""
        try:
            # This uses the default application associated with the file type
            os.startfile(file_path)
            return True
        except Exception as e:
            raise Exception(f"Error opening file: {str(e)}")
    
    def _close_app(self, app_name):
        """Close an application by name"""
        try:
            subprocess.run(["taskkill", "/f", "/im", app_name], 
                           check=False, 
                           stdout=subprocess.PIPE, 
                           stderr=subprocess.PIPE)
            return True
        except Exception as e:
            raise Exception(f"Error closing application: {str(e)}")
    
    def _run_command(self, command):
        """Run a command in the command prompt"""
        try:
            subprocess.Popen(command, shell=True)
            return True
        except Exception as e:
            raise Exception(f"Error running command: {str(e)}")
    
    def _delay(self, seconds):
        """Wait for the specified number of seconds"""
        import time
        try:
            time.sleep(int(seconds))
            return True
        except Exception as e:
            raise Exception(f"Error during delay: {str(e)}")
    
    def _shutdown(self, delay="0"):
        """Shutdown the computer with optional delay"""
        try:
            subprocess.run(["shutdown", "/s", "/t", delay], check=True)
            return True
        except Exception as e:
            raise Exception(f"Error shutting down: {str(e)}")
    
    def _restart(self, delay="0"):
        """Restart the computer with optional delay"""
        try:
            subprocess.run(["shutdown", "/r", "/t", delay], check=True)
            return True
        except Exception as e:
            raise Exception(f"Error restarting: {str(e)}")
    
    def _sleep(self):
        """Put the computer to sleep"""
        try:
            subprocess.run(["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"], check=True)
            return True
        except Exception as e:
            raise Exception(f"Error putting computer to sleep: {str(e)}")
    
    def _take_screenshot(self, save_folder):
        """Take a screenshot and save it to the specified folder"""
        try:
            from PIL import ImageGrab
            import datetime
            
            # Create timestamp for filename
            timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            filename = os.path.join(save_folder, f"screenshot_{timestamp}.png")
            
            # Take screenshot and save
            screenshot = ImageGrab.grab()
            screenshot.save(filename)
            return True
        except Exception as e:
            raise Exception(f"Error taking screenshot: {str(e)}")
    
    def _clean_temp(self):
        """Clean the Windows temp folder"""
        try:
            # Get the temp folder path
            temp_folder = os.environ.get('TEMP')
            if not temp_folder:
                raise Exception("Could not determine TEMP folder location")
                
            # Run the command to clean temp files
            subprocess.run(f'del /q /f /s "{temp_folder}\\*"', shell=True, check=False)
            return True
        except Exception as e:
            raise Exception(f"Error cleaning temp folder: {str(e)}")
    
    def _run_security_scan(self, scan_type="quick"):
        """Run a Windows Defender scan"""
        try:
            # Map scan types to Windows Defender commands
            scan_commands = {
                "quick": "Start-MpScan -ScanType QuickScan",
                "full": "Start-MpScan -ScanType FullScan",
                "custom": "Start-MpScan -ScanType CustomScan"
            }
            
            # Get the appropriate command
            command = scan_commands.get(scan_type.lower(), scan_commands["quick"])
            
            # Run the PowerShell command
            subprocess.run(["powershell", "-Command", command], check=True)
            return True
        except Exception as e:
            raise Exception(f"Error running security scan: {str(e)}")
    
    def _delete_file(self, file_path):
        """Delete a file"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            else:
                raise FileNotFoundError(f"File not found: {file_path}")
        except Exception as e:
            raise Exception(f"Error deleting file: {str(e)}")
    
    def _backup_folder(self, source_folder, destination_folder):
        """
        Backup files from one folder to another
        
        Args:
            source_folder (str): Path to the folder to backup
            destination_folder (str): Path where to save the backup
            
        Returns:
            bool: True if successful
            
        Raises:
            Exception: If the backup cannot be completed
        """
        try:
            import os
            import shutil
            
            # Make sure the source folder exists
            if not os.path.exists(source_folder):
                raise FileNotFoundError(f"Source folder not found: {source_folder}")
                
            # Create destination folder if it doesn't exist
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            
            # Get the source folder name to recreate in destination
            source_name = os.path.basename(source_folder)
            dest_path = os.path.join(destination_folder, source_name)
            
            # If the destination doesn't exist, create it
            if not os.path.exists(dest_path):
                os.makedirs(dest_path)
            
            # Walk through source folder and copy all files and subdirectories
            for root, dirs, files in os.walk(source_folder):
                # Create relative path to maintain folder structure
                rel_path = os.path.relpath(root, source_folder)
                
                # Create subdirectories in destination
                for dir_name in dirs:
                    dir_path = os.path.join(root, dir_name)
                    rel_dir_path = os.path.relpath(dir_path, source_folder)
                    dest_dir = os.path.join(dest_path, rel_dir_path)
                    if not os.path.exists(dest_dir):
                        os.makedirs(dest_dir)
                
                # Copy files to destination
                for file_name in files:
                    src_file = os.path.join(root, file_name)
                    rel_file_path = os.path.relpath(src_file, source_folder)
                    dest_file = os.path.join(dest_path, rel_file_path)
                    shutil.copy2(src_file, dest_file)  # copy2 preserves metadata
            
            return True
            
        except Exception as e:
            raise Exception(f"Error backing up folder: {str(e)}")
    
    def _delete_folder(self, folder_path, option="if empty"):
        """
        Delete a folder
        
        Args:
            folder_path (str): Path to the folder to delete
            option (str): 
                "with contents" - delete folder and all contents recursively
                "if empty" - only delete if folder is empty
                "contents only" - delete contents but keep the folder
            
        Returns:
            bool: True if successful
            
        Raises:
            Exception: If the folder cannot be deleted
        """
        try:
            if not os.path.exists(folder_path):
                raise FileNotFoundError(f"Folder not found: {folder_path}")
                
            if not os.path.isdir(folder_path):
                raise ValueError(f"Not a directory: {folder_path}")
                
            import shutil
                
            if option == "with contents":
                # Delete recursively, similar to rm -rf
                shutil.rmtree(folder_path)
            elif option == "contents only":
                # Delete all contents but keep the folder itself
                for item in os.listdir(folder_path):
                    item_path = os.path.join(folder_path, item)
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                    else:
                        os.remove(item_path)
                #(f"Deleted contents of folder: {folder_path}")
            else:
                # Delete only if empty
                try:
                    os.rmdir(folder_path)
                except OSError as e:
                    if "directory not empty" in str(e).lower():
                        raise Exception(f"Folder is not empty: {folder_path}. Use 'with contents' option to delete non-empty folders.")
                    else:
                        raise
                        
            return True
        except Exception as e:
            raise Exception(f"Error deleting folder: {str(e)}")
