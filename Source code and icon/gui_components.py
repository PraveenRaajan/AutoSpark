"""
GUI Components for the AutoSpark Application
"""

import os
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import webbrowser
from datetime import datetime
from PIL import Image, ImageTk

from task_manager import TaskManager
from task_executor import TaskExecutor
from file_handler import FileHandler

class AutomationApp:
    """Main application class for the AutoSpark application"""
    
    def __init__(self, master):
        """Initialize the application"""
        self.master = master
        self.task_manager = TaskManager()
        self.file_handler = FileHandler()
        self.task_executor = TaskExecutor()
        
        # Track if text file is currently being edited
        self.currently_editing = False
        self.current_file_path = None
        
        # Initialize timestamps for tracking unsaved changes
        self.last_save_time = 0
        self.last_modification_time = 0
        
        # Create the main UI components
        self._create_menu()
        self._create_main_interface()
        
        # Bind the window close event to check for unsaved changes
        self.master.protocol("WM_DELETE_WINDOW", self.exit_application)
        
    def _create_menu(self):
        """Create the top menu bar"""
        menubar = tk.Menu(self.master)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Task List", command=self.new_task_list)
        file_menu.add_command(label="Save Task List", command=self.save_task_list)
        file_menu.add_command(label="Save Task List As...", command=self.save_task_list_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_application)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.master.config(menu=menubar)
    
    def _configure_styles(self):
        """Configure custom styles for the application"""
        style = ttk.Style()
        
        # ---- GRID TREEVIEW STYLE ----
        # Create a custom style with visible grid lines to match the screenshot exactly
        style.configure("Grid.Treeview",
                        background="#ffffff",
                        fieldbackground="#ffffff",
                        foreground="black",
                        font=('Arial', 9),
                        rowheight=25,
                        borderwidth=1)
        
        # Make the grid lines visible by adding a border to cells
        # This creates the exact grid layout as shown in the screenshot
        style.layout("Grid.Treeview", [
            ('Grid.Treeview.treearea', {'sticky': 'nswe', 'border': '1'})
        ])
        
        # Configure the Treeview heading (column headers) to have a gray background
        # with visible borders like in the screenshot
        style.configure("Grid.Treeview.Heading",
                        background="#f0f0f0",
                        foreground="black",
                        relief="raised",
                        font=('Arial', 9, 'bold'),
                        borderwidth=1)
        
        # ---- OTHER STYLES ----
        # Configure a style for accent buttons like "Run Tasks"
        style.configure("Accent.TButton",
                        background="#0078d7",
                        foreground="white")
                        
        # Configure style for the button panel frame
        style.configure("ButtonPanel.TFrame",
                      background="#dcdad5")
    
    def _create_main_interface(self):
        """Create the main user interface components"""
        # Main frame
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Set up visual styling for the application
        self._configure_styles()
        
        # Task List Frame (Left side)
        task_frame = ttk.LabelFrame(main_frame, text="Task Sequence", padding="10")
        task_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create the task list with grid lines that match the screenshot exactly
        # Define columns for the treeview
        columns = ("id", "task_type", "task_details")
        
        # Create a container frame for treeview and scrollbar with visible border
        tree_container = tk.Frame(task_frame, bd=2, relief=tk.GROOVE, bg='#d9d9d9')
        tree_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create and configure a Treeview style with showseparators enabled
        style = ttk.Style()
        style.configure("CustomTreeview", 
                        background="white",
                        foreground="black", 
                        rowheight=25,
                        borderwidth=1, 
                        relief="solid")
        style.layout("CustomTreeview", 
                    [('CustomTreeview.treearea', {'sticky': 'nswe', 'border':'1'})])
        
        # Create the Treeview with visible borders and column separators
        self.task_tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",  # We've removed "tree" to hide the first column 
            selectmode="browse",
            style="CustomTreeview"  # Use our custom style with separators
        )
        
        # Define column headings
        self.task_tree.heading("id", text="#")
        self.task_tree.heading("task_type", text="Task Type")
        self.task_tree.heading("task_details", text="Task Details")
        
        # Match column widths to the screenshot with more visible borders
        self.task_tree.column("#0", width=30, stretch=False)  # First column for tree structure
        self.task_tree.column("id", width=40, stretch=False)
        self.task_tree.column("task_type", width=150, stretch=False)
        self.task_tree.column("task_details", width=450, stretch=True)
        
        # Set height only - ttk.Treeview doesn't support direct borderwidth configuration
        self.task_tree.configure(height=15)
        
        # Add tags for styling individual rows with alternate colors
        self.task_tree.tag_configure("evenrow", background="#f8f8f8")
        self.task_tree.tag_configure("oddrow", background="white")
        
        # Add vertical scrollbar
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        # Set up a frame with grid design
        tree_container.config(bd=1, relief=tk.SOLID, bg="#ffffff")
        
        # Grid the tree and scrollbar inside container frame
        self.task_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Configure grid weights
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)
        
        # After setting up the frame, add physical separator lines between columns
        self.create_column_separators()
        
        # Create right-click menu for the task list with appropriate options
        self.task_context_menu = tk.Menu(self.task_tree, tearoff=0)
        self.task_context_menu.add_command(label="Delete Task", command=self.delete_task)
        self.task_context_menu.add_command(label="Move Up", command=lambda: self.move_task("up"))
        self.task_context_menu.add_command(label="Move Down", command=lambda: self.move_task("down"))
        
        # Bind the right-click event
        self.task_tree.bind("<Button-3>", self.show_task_context_menu)
        
        # Task Buttons Frame with Canvas for Scrolling
        outer_frame_style = ttk.Style()
        outer_frame_style.configure("TaskPanel.TFrame", background="#dcdad5")
        outer_button_frame = ttk.Frame(main_frame, padding="5", style="TaskPanel.TFrame")
        outer_button_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add label in the outer frame
        # Configure a label style that matches the background
        label_style = ttk.Style()
        label_style.configure("TaskLabel.TLabel", background="#dcdad5")
        ttk.Label(outer_button_frame, text="Add Tasks:", style="TaskLabel.TLabel").pack(pady=(0, 5), anchor=tk.W)
        
        # Create a canvas inside the outer frame that matches the parent background
        canvas = tk.Canvas(outer_button_frame, borderwidth=0, highlightthickness=0, bg="#dcdad5")
        scrollbar = ttk.Scrollbar(outer_button_frame, orient=tk.VERTICAL, command=canvas.yview)
        
        # Create an inner frame that will hold all button groups
        button_frame = ttk.Frame(canvas, padding="5")
        button_frame.configure(style="ButtonPanel.TFrame")
        
        # Configure the canvas to work with the scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack the scrollbar and canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create a window inside the canvas with the button frame
        canvas_window = canvas.create_window((0, 0), window=button_frame, anchor="nw")
        
        # Make sure there's no border or highlight
        canvas.configure(highlightbackground="#dcdad5", highlightcolor="#dcdad5", highlightthickness=0, bd=0)
        
        # Configure canvas to resize with window
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox("all"), width=button_frame.winfo_reqwidth())
            # Make sure the window fills the width of the canvas
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())
            
        # Add mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
        def _bind_mouse_wheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
        def _unbind_mouse_wheel(event):
            canvas.unbind_all("<MouseWheel>")
            
        # Bind the mousewheel only when mouse is over the canvas
        canvas.bind("<Enter>", _bind_mouse_wheel)
        canvas.bind("<Leave>", _unbind_mouse_wheel)
        
        # Set a fixed height for the canvas to ensure scrolling works
        canvas.configure(height=400)
        
        button_frame.bind("<Configure>", configure_canvas)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
        
        # System Tasks
        system_frame = ttk.LabelFrame(button_frame, text="System Tasks", padding="5")
        system_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(system_frame, text="Open URL", command=lambda: self.add_open_url_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Open Application", command=lambda: self.add_open_app_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Open File", command=lambda: self.add_open_file_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Delete File", command=lambda: self.add_delete_file_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Delete Folder & Contents", command=lambda: self.add_delete_folder_contents_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Empty Folder", command=lambda: self.add_empty_folder_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Delete If Empty", command=lambda: self.add_delete_if_empty_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Close Application", command=lambda: self.add_close_app_task()).pack(fill=tk.X, pady=2)
        ttk.Button(system_frame, text="Run Command", command=lambda: self.add_run_command_task()).pack(fill=tk.X, pady=2)
        
        # System Control Tasks
        control_frame = ttk.LabelFrame(button_frame, text="System Control", padding="5")
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="Shutdown PC", command=lambda: self.add_shutdown_task()).pack(fill=tk.X, pady=2)
        ttk.Button(control_frame, text="Restart PC", command=lambda: self.add_restart_task()).pack(fill=tk.X, pady=2)
        ttk.Button(control_frame, text="Sleep PC", command=lambda: self.add_sleep_task()).pack(fill=tk.X, pady=2)
        
        # Utility Tasks
        utility_frame = ttk.LabelFrame(button_frame, text="Utilities", padding="5")
        utility_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(utility_frame, text="Add Delay", command=lambda: self.add_delay_task()).pack(fill=tk.X, pady=2)
        ttk.Button(utility_frame, text="Take Screenshot", command=lambda: self.add_screenshot_task()).pack(fill=tk.X, pady=2)
        ttk.Button(utility_frame, text="Clean Temp Folder", command=lambda: self.add_clean_temp_task()).pack(fill=tk.X, pady=2)
        ttk.Button(utility_frame, text="Run Security Scan", command=lambda: self.add_security_scan_task()).pack(fill=tk.X, pady=2)
        ttk.Button(utility_frame, text="Backup Folder", command=lambda: self.add_backup_folder_task()).pack(fill=tk.X, pady=2)
        
        # No Action Buttons here since they're already in the File menu
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.master, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def show_task_context_menu(self, event):
        """Show the context menu for tasks"""
        try:
            # Identify the item that was clicked on
            item_id = self.task_tree.identify_row(event.y)
            if item_id:
                # Select the clicked item
                self.task_tree.selection_set(item_id)
                # Show the context menu
                self.task_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.task_context_menu.grab_release()
    
    def new_task_list(self):
        """Create a new task list"""
        if self.task_manager.has_tasks():
            if not messagebox.askyesno("Confirm", "This will clear the current task list. Continue?"):
                return
        
        self.task_manager.clear_tasks()
        self.refresh_task_list()
        self.current_file_path = None
        
        # Reset timestamps when creating a new list
        import time
        self.last_save_time = None
        self.last_modification_time = time.time()
        
        self.status_var.set("New task list created")
    
    def open_task_list(self):
        """Open an existing task list from a .txt file"""
        file_path = filedialog.askopenfilename(
            title="Open Task List",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            tasks = self.file_handler.load_tasks(file_path)
            self.task_manager.set_tasks(tasks)
            self.refresh_task_list()
            self.current_file_path = file_path
            
            # Initialize the save time to current time (since we just loaded the file)
            import time
            self.last_save_time = time.time()
            self.last_modification_time = self.last_save_time
            
            self.status_var.set(f"Loaded tasks from {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Error Opening File", f"Could not open task list: {str(e)}")
    
    def save_task_list(self):
        """Save the current task list to a file"""
        if not self.current_file_path:
            self.save_task_list_as()
            return
            
        try:
            tasks = self.task_manager.get_tasks()
            self.file_handler.save_tasks(tasks, self.current_file_path)
            bat_file = self.file_handler.convert_to_bat(tasks, self.current_file_path)
            
            # Update the last save time to current time
            import time
            self.last_save_time = time.time()
            
            self.status_var.set(f"Tasks saved to {os.path.basename(self.current_file_path)} and {os.path.basename(bat_file)}")
        except Exception as e:
            messagebox.showerror("Error Saving File", f"Could not save task list: {str(e)}")
    
    def save_task_list_as(self):
        """
        Save the task list to a new file
        
        Returns:
            bool: True if the save was successful, False otherwise
        """
        if not self.task_manager.has_tasks():
            messagebox.showinfo("No Tasks", "No tasks to save. Add some tasks first.")
            return False
            
        file_path = filedialog.asksaveasfilename(
            title="Save Task List",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return False  # User cancelled
            
        try:
            tasks = self.task_manager.get_tasks()
            self.file_handler.save_tasks(tasks, file_path)
            bat_file = self.file_handler.convert_to_bat(tasks, file_path)
            self.current_file_path = file_path
            
            # Update the last save time to current time
            import time
            self.last_save_time = time.time()
            
            self.status_var.set(f"Tasks saved to {os.path.basename(file_path)} and {os.path.basename(bat_file)}")
            return True
        except Exception as e:
            messagebox.showerror("Error Saving File", f"Could not save task list: {str(e)}")
            return False
    
    def run_tasks(self):
        """Run the current task list"""
        if not self.task_manager.has_tasks():
            messagebox.showinfo("No Tasks", "No tasks to run. Add some tasks first.")
            return
            
        # Always save the tasks first, whether or not they've been saved before
        if self.current_file_path:
            # Save to existing file
            try:
                tasks = self.task_manager.get_tasks()
                self.file_handler.save_tasks(tasks, self.current_file_path)
                bat_file = self.file_handler.convert_to_bat(tasks, self.current_file_path)
                
                # Update the save timestamp
                import time
                self.last_save_time = time.time()
                
                self.status_var.set(f"Tasks saved to {os.path.basename(self.current_file_path)} and {os.path.basename(bat_file)}")
            except Exception as e:
                messagebox.showerror("Error Saving File", f"Could not save task list: {str(e)}")
                return
        else:
            # Need to save to a new file
            self.save_task_list_as()
            # If we still don't have a file path, the user cancelled the save dialog
            if not self.current_file_path:
                return
                
        # Execute the BAT file
        bat_path = self.file_handler.get_bat_path(self.current_file_path)
        try:
            self.task_executor.run_bat_file(bat_path)
            self.status_var.set(f"Running tasks from {os.path.basename(bat_path)}")
        except Exception as e:
            messagebox.showerror("Error Running Tasks", f"Could not run tasks: {str(e)}")
        
        def save_changes():
            """Save changes to the text file and update the BAT file"""
            try:
                # Get the content from the editor
                content = text_editor.get(1.0, tk.END)
                
                # Save to the text file - ensure file path is not None
                if self.current_file_path:
                    with open(self.current_file_path, 'w') as f:
                        f.write(content)
                    
                    # Parse the content back into tasks
                    tasks = self.file_handler.parse_text_to_tasks(content)
                    
                    # Update the task manager
                    self.task_manager.set_tasks(tasks)
                    
                    # Create a new BAT file
                    bat_file = self.file_handler.convert_to_bat(tasks, self.current_file_path)
                    bat_name = os.path.basename(bat_file) if bat_file else "Unknown"
                    
                    # Update the save timestamp
                    import time
                    self.last_save_time = time.time()
                    
                    # Update the main UI
                    self.refresh_task_list()
                    
                    status_var.set(f"Saved changes to {file_name} and {bat_name}")
                    self.status_var.set(f"Updated tasks from edited file")
                    
                    # Close the window after a short delay
                    edit_window.after(1500, edit_window.destroy)
                else:
                    # Should never get here due to initial check
                    raise ValueError("No file path available for saving")
            except Exception as e:
                messagebox.showerror("Error Saving Changes", f"Could not save changes: {str(e)}")
        
        ttk.Button(button_frame, text="Save Changes", command=save_changes).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=edit_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def create_column_separators(self):
        """Apply styling to make the Treeview grid lines visible"""
        if hasattr(self, 'task_tree'):
            # First, make sure the container frame has a visible border
            tree_container = self.task_tree.master
            if isinstance(tree_container, tk.Frame):
                tree_container.config(bd=2, relief=tk.GROOVE, bg='#e0e0e0')
                
            # Apply enhanced styling to our tree with explicit grid lines
            style = ttk.Style()
            
            # Configure main Treeview appearance with borders
            style.configure("Treeview", 
                           background="white",
                           foreground="black", 
                           rowheight=25,
                           fieldbackground="white",
                           borderwidth=1,
                           relief="solid")
            
            # Add a clear style for row highlights with borders
            style.map('Treeview', 
                     background=[('selected', '#e5f1fb')],
                     foreground=[('selected', 'black')])
            
            # Configure headings with distinct appearance and borders
            style.configure("Treeview.Heading", 
                           background="#e0e0e0",
                           foreground="black",
                           relief="raised", 
                           borderwidth=1,
                           font=('Arial', 9, 'bold'))
            
            # Apply the main Treeview style that includes borders
            self.task_tree.configure(style="Treeview")
            
            # Only show the headings, hide the tree column
            self.task_tree.config(show="headings")
            
            # Apply alternate row colors for better readability
            self.task_tree.tag_configure("evenrow", background="#f0f0f0")
            self.task_tree.tag_configure("oddrow", background="white")
            
            #("Applied enhanced Treeview styling with grid lines")
            
            # Add dark border around the container for emphasis
            # This creates a box effect that helps the table stand out
            if hasattr(tree_container, 'master') and isinstance(tree_container.master, tk.Frame):
                tree_container.master.config(padx=2, pady=2, bg='#c0c0c0')
            
    def _update_separators(self, event=None):
        """Update the column separators when the Treeview is resized"""
        # This is called when the Treeview widget changes size
        # We need to redraw our separator lines
        if hasattr(self, 'task_tree'):
            # Schedule the redraw after a short delay to ensure the Treeview
            # has fully updated its layout
            self.task_tree.after(100, self.create_column_separators)

    def refresh_task_list(self):
        """Refresh the task list display"""
        # Clear the current items
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
            
        # Add the tasks from the task manager
        for i, task in enumerate(self.task_manager.get_tasks(), 1):
            task_type = task.get("type", "")
            task_details = task.get("details", "")
            
            # Determine row tag based on even/odd index for alternating colors
            row_tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            
            # Insert task into treeview
            self.task_tree.insert("", tk.END, values=(
                i,  # Task number 
                task_type, 
                task_details
            ), tags=(row_tag,))
        
        # The style settings in main.py ensure the grid lines are visible
            
    # Using Treeview selection instead of canvas click handler
    
    def delete_task(self, task_idx=None):
        """Delete a task by index or use the currently selected one"""
        if task_idx is None:
            # If no specific task, use the selected item in the treeview
            if not self.task_manager.has_tasks():
                messagebox.showinfo("No Tasks", "No tasks to delete.")
                return
                
            # Check if an item is selected in the treeview
            selected_items = self.task_tree.selection()
            if selected_items:
                # Get the task id from the selected item
                item = selected_items[0]  # Get the first selected item
                values = self.task_tree.item(item, 'values')
                if values and len(values) > 0:
                    task_num = int(values[0])  # The task number is the first column
                    # Convert to 0-based index
                    task_idx = task_num - 1
                else:
                    # If no values, ask user for task number
                    task_num = simpledialog.askinteger("Delete Task", 
                                                    "Enter the task number to delete:",
                                                    minvalue=1, 
                                                    maxvalue=len(self.task_manager.get_tasks()))
                    if task_num is None:
                        return  # User cancelled
                    # Convert to 0-based index
                    task_idx = task_num - 1
            else:
                # No selection, ask user for task number
                task_num = simpledialog.askinteger("Delete Task", 
                                                "Enter the task number to delete:",
                                                minvalue=1, 
                                                maxvalue=len(self.task_manager.get_tasks()))
                if task_num is None:
                    return  # User cancelled
                # Convert to 0-based index
                task_idx = task_num - 1
            
        # Delete from the task manager
        result = self.task_manager.delete_task(task_idx)
        if result:
            # Update modification time
            self._update_modification_time()
            
            # Refresh the display
            self.refresh_task_list()
            self.status_var.set("Task deleted")
        else:
            messagebox.showerror("Error", "Could not delete task.")
    
    def move_task(self, direction, task_idx=None):
        """Move a task up or down in the list"""
        if task_idx is None:
            # If no specific task, use the selected item in the treeview
            if not self.task_manager.has_tasks():
                messagebox.showinfo("No Tasks", "No tasks to move.")
                return
                
            # Check if an item is selected in the treeview
            selected_items = self.task_tree.selection()
            if selected_items:
                # Get the task id from the selected item
                item = selected_items[0]  # Get the first selected item
                values = self.task_tree.item(item, 'values')
                if values and len(values) > 0:
                    task_num = int(values[0])  # The task number is the first column
                    # Convert to 0-based index
                    task_idx = task_num - 1
                else:
                    # If no values, ask user for task number
                    task_num = simpledialog.askinteger(f"Move Task {direction.capitalize()}", 
                                                    f"Enter the task number to move {direction}:",
                                                    minvalue=1, 
                                                    maxvalue=len(self.task_manager.get_tasks()))
                    if task_num is None:
                        return  # User cancelled
                    # Convert to 0-based index
                    task_idx = task_num - 1
            else:
                # No selection, ask user for task number
                task_num = simpledialog.askinteger(f"Move Task {direction.capitalize()}", 
                                                f"Enter the task number to move {direction}:",
                                                minvalue=1, 
                                                maxvalue=len(self.task_manager.get_tasks()))
                if task_num is None:
                    return  # User cancelled
                # Convert to 0-based index
                task_idx = task_num - 1
            
        # Move the task
        if direction == "up":
            success = self.task_manager.move_task_up(task_idx)
        else:  # down
            success = self.task_manager.move_task_down(task_idx)
            
        if success:
            # Update modification time
            self._update_modification_time()
            
            self.refresh_task_list()
            self.status_var.set(f"Task moved {direction}")
        else:
            messagebox.showinfo("Cannot Move", f"Cannot move task {direction} further.")
    
    def add_open_url_task(self):
        """Add a task to open a URL"""
        url = simpledialog.askstring("Open URL", "Enter the URL to open:")
        if url:
            if not (url.startswith("http://") or url.startswith("https://")):
                url = "https://" + url
            self.task_manager.add_task("open_url", url)
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Open URL {url}")
    
    def _update_modification_time(self):
        """Update the modification timestamp to mark changes"""
        import time
        self.last_modification_time = time.time()
    
    def add_open_app_task(self):
        """Add a task to open an application"""
        app_path = filedialog.askopenfilename(
            title="Select Application",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        
        if app_path:
            self.task_manager.add_task("open_app", app_path)
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Open Application {os.path.basename(app_path)}")
    
    def add_open_file_task(self):
        """Add a task to open a file with its default application"""
        file_path = filedialog.askopenfilename(
            title="Select File to Open"
        )
        
        if file_path:
            self.task_manager.add_task("open_file", file_path)
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Open File {os.path.basename(file_path)}")
    
    def add_close_app_task(self):
        """Add a task to close an application"""
        app_name = simpledialog.askstring("Close Application", 
                                          "Enter the application name to close (e.g., notepad.exe):")
        if app_name:
            self.task_manager.add_task("close_app", app_name)
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Close Application {app_name}")
    
    def add_run_command_task(self):
        """Add a task to run a command in CMD"""
        command = simpledialog.askstring("Run Command", 
                                         "Enter the command to run in command prompt:")
        if command:
            self.task_manager.add_task("run_command", command)
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Run Command {command}")
    
    def add_shutdown_task(self):
        """Add a task to shutdown the computer"""
        delay = simpledialog.askinteger("Shutdown Delay", 
                                        "Enter delay in seconds before shutdown (0 for immediate):",
                                        initialvalue=0, minvalue=0)
        if delay is not None:  # Check for Cancel
            self.task_manager.add_task("shutdown", str(delay))
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Shutdown PC with {delay}s delay")
    
    def add_restart_task(self):
        """Add a task to restart the computer"""
        delay = simpledialog.askinteger("Restart Delay", 
                                        "Enter delay in seconds before restart (0 for immediate):",
                                        initialvalue=0, minvalue=0)
        if delay is not None:  # Check for Cancel
            self.task_manager.add_task("restart", str(delay))
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Restart PC with {delay}s delay")
    
    def add_sleep_task(self):
        """Add a task to put the computer to sleep"""
        self.task_manager.add_task("sleep", "")
        self._update_modification_time()
        self.refresh_task_list()
        self.status_var.set("Added task: Sleep PC")
    
    def add_delay_task(self):
        """Add a task to wait for a specified time"""
        seconds = simpledialog.askinteger("Add Delay", 
                                         "Enter delay in seconds:",
                                         initialvalue=5, minvalue=1)
        if seconds is not None:  # Check for Cancel
            self.task_manager.add_task("delay", str(seconds))
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Delay for {seconds} seconds")
    
    def add_screenshot_task(self):
        """Add a task to take a screenshot"""
        folder = filedialog.askdirectory(title="Select folder to save screenshots")
        if folder:
            self.task_manager.add_task("screenshot", folder)
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Take screenshot to {folder}")
    
    def add_clean_temp_task(self):
        """Add a task to clean the temp folder"""
        self.task_manager.add_task("clean_temp", "")
        self._update_modification_time()
        self.refresh_task_list()
        self.status_var.set("Added task: Clean temporary files")
    
    def add_delete_file_task(self):
        """Add a task to delete a file"""
        file_path = filedialog.askopenfilename(
            title="Select File to Delete",
            filetypes=[("All Files", "*.*")]
        )
        
        if not file_path:
            return  # User cancelled
            
        self.task_manager.add_task("delete_file", file_path)
        self._update_modification_time()
        self.refresh_task_list()
        self.status_var.set(f"Added task: Delete file {os.path.basename(file_path)}")
        
    def add_delete_folder_contents_task(self):
        """Add a task to delete a folder and all its contents"""
        folder_path = filedialog.askdirectory(
            title="Select Folder to Delete with Contents"
        )
        
        if not folder_path:
            return  # User cancelled
        
        folder_name = os.path.basename(folder_path)    
        # Add the task with "with contents" option
        self.task_manager.add_task("delete_folder", folder_path, "with contents")
        self._update_modification_time()
        self.refresh_task_list()
        self.status_var.set(f"Added task: Delete folder {folder_name} and all its contents")
    
    def add_empty_folder_task(self):
        """Add a task to empty a folder but keep the folder itself"""
        folder_path = filedialog.askdirectory(
            title="Select Folder to Empty"
        )
        
        if not folder_path:
            return  # User cancelled
        
        folder_name = os.path.basename(folder_path)    
        # Add the task with "contents only" option
        self.task_manager.add_task("empty_folder", folder_path, "")
        self._update_modification_time()
        self.refresh_task_list()
        self.status_var.set(f"Added task: Empty folder {folder_name} (keep folder)")
    
    def add_delete_if_empty_task(self):
        """Add a task to delete a folder only if it's empty"""
        folder_path = filedialog.askdirectory(
            title="Select Folder to Delete (if empty)"
        )
        
        if not folder_path:
            return  # User cancelled
        
        folder_name = os.path.basename(folder_path)    
        # Add the task with "if empty" option
        self.task_manager.add_task("delete_folder_if_empty", folder_path, "")
        self._update_modification_time()
        self.refresh_task_list()
        self.status_var.set(f"Added task: Delete folder {folder_name} only if empty")
    
    def add_security_scan_task(self):
        """Add a task to run a Windows security scan"""
        scan_type = simpledialog.askstring("Security Scan", 
                                           "Enter scan type (quick or full):",
                                           initialvalue="quick")
        if scan_type:
            self.task_manager.add_task("security_scan", scan_type.lower())
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Run {scan_type} security scan")
            
    def add_backup_folder_task(self):
        """Add a task to backup a folder to another location"""
        # Create a dialog for folder selection
        backup_dialog = tk.Toplevel(self.master)
        backup_dialog.title("Backup Folder")
        backup_dialog.geometry("500x280")
        backup_dialog.resizable(False, False)
        backup_dialog.grab_set()  # Make dialog modal
        
        # Configure the dialog
        backup_dialog.columnconfigure(0, weight=1)
        backup_dialog.columnconfigure(1, weight=3)
        
        # Source folder selection
        ttk.Label(backup_dialog, text="Source Folder:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        source_var = tk.StringVar()
        ttk.Entry(backup_dialog, textvariable=source_var, width=50).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=10)
        
        def browse_source():
            folder = filedialog.askdirectory(title="Select Source Folder")
            if folder:
                source_var.set(folder)
                
        ttk.Button(backup_dialog, text="Browse...", command=browse_source).grid(row=0, column=2, padx=5, pady=10)
        
        # Destination folder selection
        ttk.Label(backup_dialog, text="Backup To:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        dest_var = tk.StringVar()
        ttk.Entry(backup_dialog, textvariable=dest_var, width=50).grid(row=1, column=1, sticky=tk.EW, padx=5, pady=10)
        
        def browse_dest():
            folder = filedialog.askdirectory(title="Select Destination Folder")
            if folder:
                dest_var.set(folder)
                
        ttk.Button(backup_dialog, text="Browse...", command=browse_dest).grid(row=1, column=2, padx=5, pady=10)
        
        # Information text
        info_text = tk.Text(backup_dialog, height=6, width=50, wrap=tk.WORD, bg="#f0f0f0")
        info_text.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky=tk.EW)
        info_text.insert(tk.END, 
            "This task will copy files from the source folder to the destination folder.\n\n"
            "If a folder with the same name already exists in the destination, files will be added to it.\n\n"
            "Any existing files with the same name will be overwritten.")
        info_text.config(state=tk.DISABLED)
        
        # Buttons
        button_frame = ttk.Frame(backup_dialog)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        def add_backup_task():
            source = source_var.get().strip()
            destination = dest_var.get().strip()
            
            if not source:
                messagebox.showwarning("Missing Source", "Please select a source folder to backup.")
                return
                
            if not destination:
                messagebox.showwarning("Missing Destination", "Please select a destination folder for the backup.")
                return
                
            # Add the task (using old-style add_task for compatibility)
            self.task_manager.add_task("backup_folder", source, destination)
            self._update_modification_time()
            self.refresh_task_list()
            self.status_var.set(f"Added task: Backup {source} to {destination}")
            backup_dialog.destroy()
            
        ttk.Button(button_frame, text="Add Backup Task", command=add_backup_task).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=backup_dialog.destroy).pack(side=tk.LEFT, padx=5)
    

    
    def not_implemented_yet(self):
        """Placeholder for features not yet implemented"""
        messagebox.showinfo("Not Implemented", "This feature is not yet implemented.")
    
    def show_about(self):
        """Show the About dialog"""
        messagebox.showinfo("About AutoSpark", 
                           "AutoSpark\n"
                           "Version 1.0\n\n"
                           "A tool for automating common Windows tasks.\n"
                           "Save your task sequences as .bat and .txt files.\n"
                           "Author:S Praveen Raajan")
                           
    def exit_application(self):
        """Exit the application, checking for unsaved changes"""
        # Check if there are unsaved changes
        if self.task_manager.has_tasks() and (not self.current_file_path or self._have_unsaved_changes()):
            response = messagebox.askyesnocancel("Unsaved Changes", 
                                               "You have unsaved changes. Would you like to save before exiting?")
            if response is None:  # User clicked Cancel
                return
            elif response:  # User clicked Yes
                self.save_task_list()
                
        # Destroy the main window to exit
        self.master.destroy()
        
    def _have_unsaved_changes(self):
        """Check if there are unsaved changes since last save"""
        # If we don't have a file path yet, but we have tasks, then we have unsaved changes
        if not self.current_file_path and self.task_manager.has_tasks():
            return True
            
        # If we have a file path, we'll implement a simple timestamp-based check
        if hasattr(self, 'last_save_time') and self.last_save_time is not None:
            # Get the modification time, defaulting to 0 if it doesn't exist
            mod_time = getattr(self, 'last_modification_time', 0) 
            # Check if we've modified tasks since the last save
            if mod_time is not None and isinstance(mod_time, (int, float)) and isinstance(self.last_save_time, (int, float)):
                return mod_time > self.last_save_time
            return False
        
        # If we don't have a last_save_time attribute yet, assume no unsaved changes
        # since the user just opened the application or loaded a file
        return False
