"""
GUI handler for the DOCX to Asana CSV Generator.

This module provides a graphical user interface for the application,
allowing users to select folders, view progress, and manage the conversion process.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Dict, Any, Optional, Callable
import threading
import queue

from .utils.logger import get_logger
from .utils.config import get_config
from .utils.validator import validate_directory_path

# Initialize logger and config
logger = get_logger(__name__)
config = get_config()


class GUIHandler:
    """
    Handler for the application's graphical user interface.
    
    Provides methods for creating and managing the GUI components,
    handling user interactions, and displaying progress and results.
    """
    
    def __init__(self, process_callback: Optional[Callable] = None):
        """
        Initialize the GUI handler.
        
        Args:
            process_callback: Optional callback function for processing files
        """
        self.root = None
        self.process_callback = process_callback
        self.input_folder_var = None
        self.output_folder_var = None
        self.progress_var = None
        self.status_var = None
        self.progress_bar = None
        self.file_listbox = None
        self.processing_thread = None
        self.queue = queue.Queue()
    
    def create_gui(self) -> tk.Tk:
        """
        Create the main GUI window and components.
        
        Returns:
            The root Tk window
        """
        # Create root window
        self.root = tk.Tk()
        self.root.title(config.get("application.name", "DOCX to Asana CSV Generator"))
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create folder selection frame
        folder_frame = ttk.LabelFrame(main_frame, text="Folder Selection", padding=10)
        folder_frame.pack(fill=tk.X, pady=5)
        
        # Input folder selection
        input_frame = ttk.Frame(folder_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="Input Folder:").pack(side=tk.LEFT, padx=5)
        self.input_folder_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_folder_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(input_frame, text="Browse...", command=self._browse_input_folder).pack(side=tk.LEFT, padx=5)
        
        # Output folder selection
        output_frame = ttk.Frame(folder_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Output Folder:").pack(side=tk.LEFT, padx=5)
        self.output_folder_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_folder_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="Browse...", command=self._browse_output_folder).pack(side=tk.LEFT, padx=5)
        
        # Create file list frame
        file_frame = ttk.LabelFrame(main_frame, text="Files to Process", padding=10)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # File listbox with scrollbar
        listbox_frame = ttk.Frame(file_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=self.file_listbox.yview)
        
        # Create progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding=10)
        progress_frame.pack(fill=tk.X, pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        status_label.pack(fill=tk.X, pady=5)
        
        # Create button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # Process button
        process_button = ttk.Button(button_frame, text="Process Files", command=self._process_files)
        process_button.pack(side=tk.RIGHT, padx=5)
        
        # Refresh button
        refresh_button = ttk.Button(button_frame, text="Refresh File List", command=self._refresh_file_list)
        refresh_button.pack(side=tk.RIGHT, padx=5)
        
        # Set up queue processing
        self.root.after(100, self._process_queue)
        
        return self.root
    
    def _browse_input_folder(self) -> None:
        """Open a dialog to select the input folder."""
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_folder_var.set(folder)
            self._refresh_file_list()
    
    def _browse_output_folder(self) -> None:
        """Open a dialog to select the output folder."""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder_var.set(folder)
    
    def _refresh_file_list(self) -> None:
        """Refresh the list of DOCX files in the input folder."""
        input_folder = self.input_folder_var.get()
        if not input_folder or not validate_directory_path(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder")
            return
        
        # Clear the listbox
        self.file_listbox.delete(0, tk.END)
        
        # Find all DOCX files in the input folder
        docx_files = []
        try:
            for file in os.listdir(input_folder):
                if file.lower().endswith(".docx") and not file.startswith("~$"):
                    docx_files.append(file)
        except Exception as e:
            logger.error(f"Error listing files in folder: {e}")
            messagebox.showerror("Error", f"Error listing files: {e}")
            return
        
        # Add files to the listbox
        for file in sorted(docx_files):
            self.file_listbox.insert(tk.END, file)
        
        # Update status
        self.status_var.set(f"Found {len(docx_files)} DOCX files")
    
    def _process_files(self) -> None:
        """Process the selected DOCX files."""
        # Check if already processing
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showinfo("Info", "Processing is already in progress")
            return
        
        # Validate input folder
        input_folder = self.input_folder_var.get()
        if not input_folder or not validate_directory_path(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder")
            return
        
        # Validate output folder
        output_folder = self.output_folder_var.get()
        if not output_folder:
            messagebox.showerror("Error", "Please select an output folder")
            return
        
        # Create output folder if it doesn't exist
        if not validate_directory_path(output_folder, create_if_missing=True):
            messagebox.showerror("Error", "Could not create output folder")
            return
        
        # Get selected files or all files if none selected
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            selected_files = [self.file_listbox.get(i) for i in selected_indices]
        else:
            selected_files = [self.file_listbox.get(i) for i in range(self.file_listbox.size())]
        
        if not selected_files:
            messagebox.showinfo("Info", "No files to process")
            return
        
        # Create full file paths
        file_paths = [os.path.join(input_folder, file) for file in selected_files]
        
        # Reset progress
        self.progress_var.set(0)
        self.status_var.set("Processing files...")
        
        # Start processing thread
        self.processing_thread = threading.Thread(
            target=self._run_processing,
            args=(file_paths, output_folder)
        )
        self.processing_thread.daemon = True
        self.processing_thread.start()
    
    def _run_processing(self, file_paths: List[str], output_folder: str) -> None:
        """
        Run the file processing in a separate thread.
        
        Args:
            file_paths: List of file paths to process
            output_folder: Output folder for CSV files
        """
        try:
            if self.process_callback:
                # Process files with progress updates
                total_files = len(file_paths)
                for i, file_path in enumerate(file_paths):
                    # Update progress
                    progress = (i / total_files) * 100
                    self.queue.put(("progress", progress))
                    self.queue.put(("status", f"Processing {i+1}/{total_files}: {os.path.basename(file_path)}"))
                    
                    # Process the file
                    try:
                        result = self.process_callback(file_path, output_folder)
                        status = "Success" if result else "Failed"
                        self.queue.put(("file_status", (file_path, status)))
                    except Exception as e:
                        logger.error(f"Error processing file {file_path}: {e}")
                        self.queue.put(("file_status", (file_path, f"Error: {e}")))
                
                # Update final progress
                self.queue.put(("progress", 100))
                self.queue.put(("status", f"Completed processing {total_files} files"))
                self.queue.put(("complete", None))
            else:
                self.queue.put(("status", "No processing callback defined"))
                self.queue.put(("complete", None))
        except Exception as e:
            logger.error(f"Error in processing thread: {e}")
            self.queue.put(("status", f"Error: {e}"))
            self.queue.put(("complete", None))
    
    def _process_queue(self) -> None:
        """Process messages from the queue."""
        try:
            while True:
                message_type, data = self.queue.get_nowait()
                
                if message_type == "progress":
                    self.progress_var.set(data)
                elif message_type == "status":
                    self.status_var.set(data)
                elif message_type == "file_status":
                    file_path, status = data
                    logger.info(f"File {file_path}: {status}")
                elif message_type == "complete":
                    messagebox.showinfo("Complete", "Processing completed")
                
                self.queue.task_done()
        except queue.Empty:
            pass
        
        # Schedule next queue check
        self.root.after(100, self._process_queue)
    
    def run(self) -> None:
        """Run the GUI main loop."""
        if not self.root:
            self.create_gui()
        
        self.root.mainloop()


def create_gui(process_callback: Optional[Callable] = None) -> GUIHandler:
    """
    Create a GUI handler with the specified process callback.
    
    Args:
        process_callback: Callback function for processing files
    
    Returns:
        A GUIHandler instance
    """
    handler = GUIHandler(process_callback)
    handler.create_gui()
    return handler


def run_gui(process_callback: Optional[Callable] = None) -> None:
    """
    Create and run the GUI with the specified process callback.
    
    Args:
        process_callback: Callback function for processing files
    """
    handler = create_gui(process_callback)
    handler.run()
