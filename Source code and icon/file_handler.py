"""
File Handler for the AutoSpark Application
Handles saving and loading tasks to/from files
"""
import os
import json
import time

class FileHandler:
    """Handles file operations for the task automation application"""
    
    def __init__(self):
        """Initialize the file handler"""
        pass
        
    def save_tasks(self, tasks, file_path):
        """
        Save tasks to a text file
        
        Args:
            tasks (list): List of task dictionaries
            file_path (str): Path to save the file
            
        Returns:
            str: Path to the saved file
            
        Raises:
            Exception: If the file cannot be saved
        """
        try:
            # If the file_path doesn't end with .txt, add it
            if not file_path.lower().endswith('.txt'):
                file_path += '.txt'
            
            with open(file_path, 'w') as f:
                # Write header
                f.write("# AutoSpark Application - Task List\n")
                f.write("# Generated on: " + time.strftime("%Y-%m-%d %H:%M:%S") + "\n")
                f.write("# Format: [Task Type] | [Details] | [Additional Info (optional)]\n\n")
                
                # Write task entries
                for task in tasks:
                    task_type = task.get("type", "")
                    details = task.get("details", "")
                    additional = task.get("additional", "")
                    
                    f.write(f"[{task_type}] | {details}")
                    if additional:
                        f.write(f" | {additional}")
                    f.write("\n")
            
            # Convert to BAT file also
            self.convert_to_bat(tasks, file_path)
            
            return file_path
            
        except Exception as e:
            raise Exception(f"Error saving tasks to file: {str(e)}")
            
    def load_tasks(self, file_path):
        """
        Load tasks from a text file
        
        Args:
            file_path (str): Path to the file to load
            
        Returns:
            list: List of task dictionaries
            
        Raises:
            Exception: If the file cannot be loaded
        """
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
                
            with open(file_path, 'r') as f:
                content = f.read()
                
            tasks = self.parse_text_to_tasks(content)
            return tasks
            
        except Exception as e:
            raise Exception(f"Error loading tasks from file: {str(e)}")
            
    def parse_text_to_tasks(self, content):
        """
        Parse text content into task dictionaries
        
        Args:
            content (str): Text content to parse
            
        Returns:
            list: List of task dictionaries
        """
        tasks = []
        lines = content.strip().split('\n')
        
        # Skip comments and empty lines
        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
                
            # Parse task format: [task_type] | details | additional
            if line.startswith('[') and ']' in line and '|' in line:
                # Extract task type from [...]
                task_type_end = line.find(']')
                task_type = line[1:task_type_end].strip()
                
                # Split the rest by pipe symbol
                rest = line[task_type_end + 1:].strip()
                parts = [p.strip() for p in rest.split('|', 1)]
                
                details = parts[0] if parts else ""
                
                # Check if there's additional information
                additional = ""
                if len(parts) > 1 and '|' in parts[1]:
                    additional_parts = parts[1].split('|', 1)
                    additional = additional_parts[0].strip()
                elif len(parts) > 1:
                    additional = parts[1].strip()
                
                tasks.append({
                    "type": task_type,
                    "details": details,
                    "additional": additional
                })
                
        return tasks
        
    def convert_to_bat(self, tasks, txt_file_path):
        """
        Convert tasks to a .bat file
        
        Args:
            tasks (list): List of task dictionaries
            txt_file_path (str): Path to the original .txt file
            
        Returns:
            str: Path to the created .bat file
            
        Raises:
            Exception: If the file cannot be created
        """
        try:
            # Get the corresponding .bat file path
            bat_file_path = self.get_bat_path(txt_file_path)
            
            with open(bat_file_path, 'w') as f:
                # Write batch file header
                f.write("@echo off\n")
                f.write("echo AutoSpark Application - Task Execution\n")
                f.write("echo Generated from: " + os.path.basename(txt_file_path) + "\n")
                f.write("echo Run time: %DATE% %TIME%\n")
                f.write("echo.\n\n")
                
                for i, task in enumerate(tasks, 1):
                    task_type = task.get("type", "")
                    details = task.get("details", "")
                    additional = task.get("additional", "")
                    
                    f.write(f"REM Task {i}: {task_type}\n")
                    f.write(f"echo Executing Task {i}: {task_type}\n")
                    
                    if task_type == "open_url":
                        self._write_url_code(f, details)
                    elif task_type == "open_app":
                        f.write(f'start "" "{details}"\n')
                    elif task_type == "open_file":
                        f.write(f'start "" "{details}"\n')
                    elif task_type == "close_app":
                        f.write(f'taskkill /f /im "{details}" >nul 2>&1\n')
                        f.write('if %ERRORLEVEL% EQU 0 (echo Successfully closed {details}) else (echo Failed to close {details})\n')
                    elif task_type == "run_command":
                        f.write(f'{details}\n')
                    elif task_type == "delay":
                        seconds = int(details)
                        f.write(f'echo Waiting for {seconds} seconds...\n')
                        f.write(f'timeout /t {seconds} /nobreak >nul\n')
                    elif task_type == "shutdown":
                        f.write(f'echo System will shutdown in {details} seconds...\n')
                        f.write(f'shutdown /s /t {details}\n')
                    elif task_type == "restart":
                        f.write(f'echo System will restart in {details} seconds...\n')
                        f.write(f'shutdown /r /t {details}\n')
                    elif task_type == "sleep":
                        f.write('echo Putting system to sleep...\n')
                        f.write('rundll32.exe powrprof.dll,SetSuspendState 0,1,0\n')
                    elif task_type == "screenshot":
                        f.write('echo Taking screenshot...\n')

                        # Ensure the folder exists
                        safe_path = details.replace('/', '\\')
                        f.write(f'if not exist "{safe_path}" mkdir "{safe_path}"\n')

                        # Safe PowerShell script — avoids string expansion issues
                        f.write(f'powershell -NoProfile -ExecutionPolicy Bypass -Command "$ts=Get-Date -Format \\"yyyy-MM-dd_HH-mm-ss\\"; ')
                        f.write(f'$path=\\"{safe_path}\\screenshot_$ts.png\\"; ')
                        f.write('[void][Reflection.Assembly]::LoadWithPartialName(\\"System.Windows.Forms\\"); ')
                        f.write('[void][Reflection.Assembly]::LoadWithPartialName(\\"System.Drawing\\"); ')
                        f.write('$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds; ')
                        f.write('$bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height; ')
                        f.write('$g = [System.Drawing.Graphics]::FromImage($bmp); ')
                        f.write('$g.CopyFromScreen($bounds.X, $bounds.Y, 0, 0, $bounds.Size); ')
                        f.write('$bmp.Save($path); ')
                        f.write('$g.Dispose(); $bmp.Dispose(); ')
                        f.write('Write-Host \\"Screenshot saved to: $path\\""\n')

                        f.write('if %ERRORLEVEL% NEQ 0 echo ERROR: Failed to take screenshot\n\n')

                    elif task_type == "clean_temp":
                        f.write('echo Cleaning temporary files...\n')
                        f.write('del /q /f /s "%TEMP%\\*" >nul 2>&1\n')
                        f.write('echo Temporary files cleaned.\n')
                    elif task_type == "security_scan":
                        scan_type = details.lower()
                        ps_command = ""
                        if scan_type == "quick":
                            ps_command = "Start-MpScan -ScanType QuickScan"
                        elif scan_type == "full":
                            ps_command = "Start-MpScan -ScanType FullScan"
                        else:
                            ps_command = "Start-MpScan -ScanType CustomScan"
                            
                        f.write(f'echo Running {scan_type} security scan...\n')
                        f.write(f'powershell -Command "{ps_command}"\n')
                    elif task_type == "delete_file":
                        # Make sure the path uses proper Windows path format with double backslashes
                        safe_path = details.replace('/', '\\')
                        
                        f.write(f'echo Deleting file: {safe_path}\n')
                        f.write(f'if exist "{safe_path}" (\n')
                        # Use DEL command with full path in quotes, both /F (force) and /Q (quiet) flags
                        f.write(f'  del /F /Q "{safe_path}"\n')
                        f.write(f'  echo Return code: %ERRORLEVEL%\n') 
                        f.write(f'  if %ERRORLEVEL% NEQ 0 (\n')
                        f.write(f'    echo ERROR: Failed to delete file with code %ERRORLEVEL%\n')
                        f.write(f'  ) else (\n')
                        f.write(f'    echo Successfully deleted file: {safe_path}\n')
                        f.write(f'  )\n')
                        f.write(f') else (\n')
                        f.write(f'  echo WARNING: File not found: {safe_path}\n')
                        f.write(f')\n')
                    elif task_type == "empty_folder":
                            # Ensure the path uses proper Windows path format with double backslashes
                        safe_path = details.replace('/', '\\')
                        f.write(f'echo Emptying folder: {safe_path}\n')
                        f.write(f'if exist "{safe_path}" (\n')
                        f.write(f'  echo Deleting all subfolders in: {safe_path}\n')
                        f.write(f'  for /d %%i in ("{safe_path}\\*") do (\n')
                        f.write(f'    rmdir /s /q "%%i" >nul 2>&1\n')
                        f.write(f'  )\n')
                        f.write(f'  echo Deleting all files in: {safe_path}\n')
                        f.write(f'  del /f /q "{safe_path}\\*" >nul 2>&1\n')
                        f.write(f'  echo Return code: %ERRORLEVEL%\n')
                        f.write(f'  if %ERRORLEVEL% NEQ 0 (\n')
                        f.write(f'    echo ERROR: Failed to empty folder with code %ERRORLEVEL%\n')
                        f.write(f'  ) else (\n')
                        f.write(f'    echo Successfully emptied folder: {safe_path}\n')
                        f.write(f'  )\n')
                        f.write(f') else (\n')
                        f.write(f'  echo WARNING: Folder not found: {safe_path}\n')
                        f.write(f')\n')
                    elif task_type == "delete_folder":
                        # Make sure the path uses proper Windows path format with double backslashes
                        safe_path = details.replace('/', '\\')
                        
                        f.write(f'echo Deleting folder: {safe_path}\n')
                        # Delete recursively
                        f.write(f'if exist "{safe_path}" (\n')
                        f.write(f'  rmdir /s /q "{safe_path}"\n')
                        f.write(f'  echo Return code: %ERRORLEVEL%\n')
                        f.write(f'  if %ERRORLEVEL% NEQ 0 (\n')
                        f.write(f'    echo ERROR: Failed to delete folder with code %ERRORLEVEL%\n')
                        f.write(f'  ) else (\n')
                        f.write(f'    echo Successfully deleted folder and its contents: {safe_path}\n')
                        f.write(f'  )\n')
                        f.write(f') else (\n')
                        f.write(f'  echo WARNING: Folder not found: {safe_path}\n')
                        f.write(f')\n')
                    elif task_type == "delete_folder_if_empty":
                        safe_path = details.replace('/', '\\')
                        f.write(f'echo Attempting to delete folder (only if empty): {safe_path}\n')
                        f.write(f'if exist "{safe_path}" (\n')
                        f.write(f'  rmdir /q "{safe_path}"\n')
                        f.write(f'  if exist "{safe_path}" (\n')
                        f.write(f'    echo WARNING: Could not delete folder {safe_path} — it may not be empty or is locked\n')
                        f.write(f'  ) else (\n')
                        f.write(f'    echo Successfully deleted empty folder: {safe_path}\n')
                        f.write(f'  )\n')
                        f.write(f') else (\n')
                        f.write(f'  echo WARNING: Folder not found: {safe_path}\n')
                        f.write(f')\n')
                    elif task_type == "backup_folder":
                        # Make sure the paths use proper Windows path format with double backslashes
                        safe_source = details.replace('/', '\\')
                        safe_dest = additional.replace('/', '\\')
                        
                        f.write(f'echo Backing up folder: {safe_source}\n')
                        f.write(f'echo to: {safe_dest}\n')
                        
                        # Check if source exists
                        f.write(f'if not exist "{safe_source}" (\n')
                        f.write(f'  echo ERROR: Source folder not found: {safe_source}\n')
                        f.write(f'  goto :backup_error\n')
                        f.write(f')\n')
                        
                        # Create destination folder if it doesn't exist
                        f.write(f'if not exist "{safe_dest}" (\n')
                        f.write(f'  mkdir "{safe_dest}"\n')
                        f.write(f'  if %ERRORLEVEL% NEQ 0 (\n')
                        f.write(f'    echo ERROR: Could not create destination folder: {safe_dest}\n')
                        f.write(f'    goto :backup_error\n')
                        f.write(f'  )\n')
                        f.write(f')\n')
                        
                        # Get the folder name from the path
                        f.write(f'for %%I in ("{safe_source}") do set "source_name=%%~nxI"\n')
                        
                        # Create destination folder if it doesn't exist
                        f.write(f'set "dest_path={safe_dest}\\%source_name%"\n')
                        f.write(f'if not exist "%dest_path%" (\n')
                        f.write(f'  mkdir "%dest_path%"\n')
                        f.write(f'  if %ERRORLEVEL% NEQ 0 (\n')
                        f.write(f'    echo ERROR: Could not create destination folder: %dest_path%\n')
                        f.write(f'    goto :backup_error\n')
                        f.write(f'  )\n')
                        f.write(f')\n')
                        
                        # Copy the folder contents with xcopy (overwriting existing files)
                        f.write(f'xcopy "{safe_source}\\*" "%dest_path%" /E /H /C /I /Y\n')
                        f.write(f'if %ERRORLEVEL% NEQ 0 (\n')
                        f.write(f'  echo ERROR: Backup operation failed\n')
                        f.write(f'  goto :backup_error\n')
                        f.write(f') else (\n')
                        f.write(f'  echo Successfully backed up {safe_source} to %dest_path%\n')
                        f.write(f')\n')
                        f.write(f'goto :backup_end\n')
                        f.write(f':backup_error\n')
                        f.write(f'echo Backup operation failed\n')
                        f.write(f':backup_end\n')
                    else:
                        f.write(f'echo Unknown task type: {task_type}\n')
                    
                    f.write('echo.\n\n')
                
                # Add a pause at the end to keep the window open
                f.write('echo All tasks completed.\n')
                f.write('pause\n')
            
            return bat_file_path
            
        except Exception as e:
            raise Exception(f"Error creating BAT file: {str(e)}")
    
    def get_bat_path(self, txt_file_path):
        """
        Get the corresponding .bat file path for a .txt file path
        
        Args:
            txt_file_path (str): Path to the .txt file
            
        Returns:
            str: Path to the corresponding .bat file
        """
        # Replace .txt extension with .bat, or add .bat if no extension
        base, ext = os.path.splitext(txt_file_path)
        return base + '.bat'
    
    def _write_url_code(self, file, url):
        """
        Write batch code to open a URL using the default web browser.
        Uses the 'start' command, which is fast and compatible with most Windows versions.

        Args:
            file: An open file object to write to
            url (str): The URL to open
        """
        file.write('echo Opening URL in default web browser...\n')

        # Ensure the URL has a proper protocol
        if not (url.startswith('http://') or url.startswith('https://') or 
                url.startswith('ftp://') or url.startswith('file://')):
            file.write('echo No protocol specified, using https:// by default\n')
            url = 'https://' + url

        # Use the 'start' command to open the URL
        file.write('echo Using start command...\n')
        file.write(f'start "" "{url}"\n')
        file.write('if %ERRORLEVEL% NEQ 0 (\n')
        file.write('    echo ERROR: Failed to open URL.\n')
        file.write(f'    echo URL: {url}\n')
        file.write(') else (\n')
        file.write(f'    echo Successfully opened URL: {url}\n')
        file.write(')\n\n')

        
        # End label
        file.write(':url_end\n')