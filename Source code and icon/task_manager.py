"""
Task Manager for the AutoSpark Application
Handles the storage and manipulation of automation tasks
"""

class TaskManager:
    """Manages the collection of automation tasks"""
    
    def __init__(self):
        """Initialize with an empty task list"""
        self._tasks = []
    
    def add_task(self, task_type, details, additional=""):
        """
        Add a new task to the list
        
        Args:
            task_type (str): The type of task (e.g., open_url, delay)
            details (str): The main details of the task
            additional (str, optional): Additional information or parameters
        """
        self._tasks.append({
            "type": task_type,
            "details": details,
            "additional": additional
        })
    
    def get_tasks(self):
        """Return the list of tasks"""
        return self._tasks.copy()
    
    def set_tasks(self, tasks):
        """
        Replace the current task list with a new one
        
        Args:
            tasks (list): List of task dictionaries
        """
        self._tasks = tasks.copy()
    
    def delete_task(self, index):
        """
        Delete a task by index
        
        Args:
            index (int): The index of the task to delete
            
        Returns:
            bool: True if successful, False otherwise
        """
        if 0 <= index < len(self._tasks):
            del self._tasks[index]
            return True
        return False
    
    def move_task_up(self, index):
        """
        Move a task up in the list (swap with previous task)
        
        Args:
            index (int): The index of the task to move
            
        Returns:
            bool: True if successful, False otherwise
        """
        if 0 < index < len(self._tasks):
            self._tasks[index], self._tasks[index-1] = self._tasks[index-1], self._tasks[index]
            return True
        return False
    
    def move_task_down(self, index):
        """
        Move a task down in the list (swap with next task)
        
        Args:
            index (int): The index of the task to move
            
        Returns:
            bool: True if successful, False otherwise
        """
        if 0 <= index < len(self._tasks) - 1:
            self._tasks[index], self._tasks[index+1] = self._tasks[index+1], self._tasks[index]
            return True
        return False
    
    def clear_tasks(self):
        """Clear all tasks from the list"""
        self._tasks = []
    
    def has_tasks(self):
        """Check if there are any tasks in the list"""
        return len(self._tasks) > 0
    
    def get_task_by_index(self, index):
        """
        Get a task by its index
        
        Args:
            index (int): The index of the task
            
        Returns:
            dict: The task dictionary, or None if index is invalid
        """
        if 0 <= index < len(self._tasks):
            return self._tasks[index].copy()
        return None
        

