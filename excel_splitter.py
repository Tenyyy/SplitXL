"""
SplitXL

Description:
    A GUI-based utility for splitting a large Excel file into smaller, more
    manageable chunks while preserving cell formatting, styles, and formulas.
    The application runs the splitting process in a separate thread to ensure
    the user interface remains responsive, providing real-time progress
    updates and a cancellation option.

The program executes the following numbered steps:
  1. Launches a GUI and prompts the user to select an input `.xlsx` file.
  2. Asks the user for an output directory, defaulting to the input file's location.
  3. Prompts for configuration: number of data rows per chunk and number of
     header rows to repeat in each new file.
  4. Presents a choice to preserve formulas (may cause #REF! errors in the
     split files) or to save only the static, calculated values.
  5. Performs a pre-scan of the input file to determine the total number of
     chunks that will be created for an accurate progress bar.
  6. Launches the main splitting operation in a separate, non-blocking thread
     to keep the GUI responsive.
  7. Displays a progress window showing the current status and a progress bar,
     complete with a "Cancel" button. The terminal output mirrors this progress.
  8. The worker thread iterates through each chunk, copying header rows, data
     rows, and preserving all cell formatting (styles, comments, merged cells, etc.).
  9. Checks for the cancellation signal before processing each new chunk to
     allow for a graceful exit.
 10. Saves each completed chunk as a new `.xlsx` file in the chosen output directory.
 11. Upon completion, error, or cancellation, displays a final summary report
     in both a GUI dialog box and the terminal.

Usage:
    - Ensure required libraries are installed:
          pip install openpyxl
    - Run the script from a terminal or by executing the file directly:
          python excel_splitter.py
          
Author:     Vitalii Starosta
GitHub:     https://github.com/sztaroszta
License:    GNU Affero General Public License v3 (AGPLv3)
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
from copy import copy
import openpyxl
from openpyxl.utils import get_column_letter
import threading
import queue

# Define the maximum number of rows in an Excel sheet for validation
MAX_ROW = 1048576

class ProgressManager:
    """Manages the Toplevel progress window UI."""

    def __init__(self, parent, title, total_steps, cancel_event):
        """
        Initializes the ProgressManager window.

        Args:
            parent (tk.Tk): The root Tkinter window.
            title (str): The title for the progress window.
            total_steps (int): The total number of steps for the progress bar.
            cancel_event (threading.Event): The event to set when cancellation is requested.
        """
        self.parent = parent
        self.total_steps = total_steps
        self.cancel_event = cancel_event

        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.resizable(False, False)
        self.window.protocol("WM_DELETE_WINDOW", self.request_cancel)

        self.status_label = tk.Label(self.window, text="Initializing...", padx=20, pady=10, width=50)
        self.status_label.pack()

        self.progress_bar = ttk.Progressbar(self.window, orient="horizontal", length=350, mode="determinate", maximum=total_steps)
        self.progress_bar.pack(padx=20, pady=5)

        self.cancel_button = tk.Button(self.window, text="Cancel", command=self.request_cancel, width=10)
        self.cancel_button.pack(pady=10)

        self.parent.update_idletasks()
        
    def update(self, current_step, status_text):
        """
        Updates the GUI and terminal progress indicators.

        Args:
            current_step (int): The current step in the process.
            status_text (str): The text to display as the current status.
        """
        self.progress_bar['value'] = current_step
        self.status_label.config(text=status_text)
        
        progress_percent = (current_step / self.total_steps) * 100
        bar_length = 30
        filled_length = int(bar_length * current_step // self.total_steps)
        bar = '█' * filled_length + '-' * (bar_length - filled_length)
        terminal_text = f"\rProgress: |{bar}| {progress_percent:.1f}% ({current_step}/{self.total_steps}) - Processing..."
        sys.stdout.write(terminal_text)
        sys.stdout.flush()
        
        self.parent.update_idletasks()

    def request_cancel(self):
        """Flags the operation for cancellation by setting the threading event."""
        self.status_label.config(text="Cancellation requested...")
        self.cancel_event.set()

    def close(self):
        """Closes the progress window and prints a final newline to the terminal."""
        sys.stdout.write('\n')
        self.window.destroy()

def _copy_cell_properties(source_cell, target_cell):
    """
    Copies value, style, hyperlink, and comment from a source cell to a target cell.

    Args:
        source_cell (openpyxl.cell.Cell): The cell to copy from.
        target_cell (openpyxl.cell.Cell): The cell to copy to.
    """
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    if source_cell.hyperlink:
        target_cell.hyperlink = copy(source_cell.hyperlink)
    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)

def _copy_row_formatting(ws_source, ws_target, source_row_idx, target_row_idx, max_col):
    """
    Copies all cell properties and row height for an entire row.

    Args:
        ws_source (openpyxl.worksheet.Worksheet): The source worksheet.
        ws_target (openpyxl.worksheet.Worksheet): The target worksheet.
        source_row_idx (int): The index of the row to copy from.
        target_row_idx (int): The index of the row to copy to.
        max_col (int): The maximum number of columns to copy.
    """
    for col_idx in range(1, max_col + 1):
        source_cell = ws_source.cell(row=source_row_idx, column=col_idx)
        target_cell = ws_target.cell(row=target_row_idx, column=col_idx)
        _copy_cell_properties(source_cell, target_cell)
    if source_row_idx in ws_source.row_dimensions:
        source_rd = ws_source.row_dimensions[source_row_idx]
        target_rd = ws_target.row_dimensions[target_row_idx]
        target_rd.height = source_rd.height

def _copy_merged_cells(ws_source, ws_target, min_row, max_row):
    """
    Copies merged cell ranges from a source worksheet that are fully
    contained within a given row range.

    Args:
        ws_source (openpyxl.worksheet.Worksheet): The source worksheet.
        ws_target (openpyxl.worksheet.Worksheet): The target worksheet.
        min_row (int): The minimum row index of the range to check for merges.
        max_row (int): The maximum row index of the range to check for merges.
    """
    for merged_range in ws_source.merged_cells.ranges:
        if merged_range.min_row >= min_row and merged_range.max_row <= max_row:
            try:
                ws_target.merge_cells(str(merged_range))
            except Exception as e:
                print(f"Warning: Could not merge range {merged_range}: {e}")

def split_excel_file_with_formatting(input_file, output_directory, chunk_size, heading_rows, preserve_formulas, progress_queue, cancel_event):
    """
    Performs the Excel splitting in a worker thread. Communicates progress,
    results, and errors back to the main thread via a queue.

    Args:
        input_file (str): Path to the source Excel file.
        output_directory (str): Path to the folder to save chunked files.
        chunk_size (int): Number of data rows per output file.
        heading_rows (int): Number of header rows to repeat in each file.
        preserve_formulas (bool): If True, preserve formulas; otherwise, save calculated values.
        progress_queue (queue.Queue): The queue for sending status updates to the main thread.
        cancel_event (threading.Event): The event to check for cancellation requests.
    """
    input_filename_base = os.path.splitext(os.path.basename(input_file))[0]
    files_created = 0

    try:
        load_data_only = not preserve_formulas
        wb_source = openpyxl.load_workbook(input_file, data_only=load_data_only, read_only=False)
        ws_source = wb_source.active
    except Exception as e:
        result = {'status': 'error', 'message': f"Error loading Excel file: {e}", 'files_created': 0}
        progress_queue.put({'type': 'result', 'data': result})
        return

    total_rows, max_col = ws_source.max_row, ws_source.max_column

    if total_rows == 0 or max_col == 0:
        result = {'status': 'success', 'message': "Input file's active sheet was empty.", 'files_created': 0}
        progress_queue.put({'type': 'result', 'data': result})
        return
    
    heading_rows = max(0, min(heading_rows, total_rows))
    data_rows_to_process = total_rows - heading_rows

    if data_rows_to_process <= 0:
        result = {'status': 'success', 'message': 'All rows were headers. One file created.', 'files_created': 1}
        progress_queue.put({'type': 'result', 'data': result})
        return

    num_chunks = (data_rows_to_process + chunk_size - 1) // chunk_size

    for i in range(num_chunks):
        if cancel_event.is_set():
            result = {'status': 'cancelled', 'message': 'Operation cancelled.', 'files_created': files_created}
            progress_queue.put({'type': 'result', 'data': result})
            return

        source_data_start_row = heading_rows + (i * chunk_size) + 1
        source_data_end_row = min(heading_rows + ((i + 1) * chunk_size), total_rows)
        
        status_text = f"Processing chunk {i+1}/{num_chunks} (Rows {source_data_start_row}-{source_data_end_row})"
        progress_queue.put({'type': 'progress', 'step': i + 1, 'status': status_text})

        wb_chunk = openpyxl.Workbook()
        ws_chunk = wb_chunk.active
        ws_chunk.title = ws_source.title

        if ws_source.auto_filter.ref:
            ws_chunk.auto_filter.ref = ws_source.auto_filter.ref

        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            if col_letter in ws_source.column_dimensions:
                ws_chunk.column_dimensions[col_letter].width = ws_source.column_dimensions[col_letter].width

        current_target_row = 1
        if heading_rows > 0:
            for r_idx in range(1, heading_rows + 1):
                _copy_row_formatting(ws_source, ws_chunk, r_idx, current_target_row, max_col)
                current_target_row += 1
            _copy_merged_cells(ws_source, ws_chunk, 1, heading_rows)

        for source_r_idx in range(source_data_start_row, source_data_end_row + 1):
            _copy_row_formatting(ws_source, ws_chunk, source_r_idx, current_target_row, max_col)
            current_target_row += 1

        _copy_merged_cells(ws_source, ws_chunk, source_data_start_row, source_data_end_row)

        output_file_name = f"{input_filename_base}_rows_{source_data_start_row}-{source_data_end_row}.xlsx"
        output_path = os.path.join(output_directory, output_file_name)
        try:
            wb_chunk.save(output_path)
            files_created += 1
        except Exception as e:
            result = {'status': 'error', 'message': f"Error saving {output_path}: {e}", 'files_created': files_created}
            progress_queue.put({'type': 'result', 'data': result})
            return
    
    result = {'status': 'success', 'message': f'Successfully created {files_created} files.', 'files_created': files_created}
    progress_queue.put({'type': 'result', 'data': result})

class App:
    """The main application class that orchestrates the GUI, user input, and the background worker thread."""
    def __init__(self, root):
        """
        Initializes and runs the main application.

        Args:
            root (tk.Tk): The root Tkinter window, which will be withdrawn.
        """
        self.root = root
        self.progress_manager = None
        self.run()

    def get_user_input(self):
        """
        Handles all initial user dialogs to get processing parameters.

        Returns:
            bool: True if the user completed all prompts, False if they cancelled at any point.
        """
        self.input_file = filedialog.askopenfilename(title="Select Input Excel File (.xlsx)", filetypes=[("Excel files", "*.xlsx")])
        if not self.input_file: return False
        
        self.output_directory = filedialog.askdirectory(title="Select Output Directory", initialdir=os.path.dirname(self.input_file))
        if not self.output_directory: return False

        try:
            self.chunk_size = simpledialog.askinteger("Data Rows Per File", "Enter data rows per file (excluding headers):", initialvalue=5000, minvalue=1)
            if self.chunk_size is None: return False
            self.heading_rows = simpledialog.askinteger("Header Rows", "Enter header rows to repeat in each file:", initialvalue=1, minvalue=0)
            if self.heading_rows is None: return False
        except (TypeError, ValueError):
            return False

        self.preserve_formulas = messagebox.askyesno(
            title="Preserve Formulas?",
            message="Do you want to preserve formulas?\n\n"
                    "• 'Yes' will keep formulas (e.g., '=A1+B1'), but may cause #REF! errors in new files.\n"
                    "• 'No' will copy only the calculated values (e.g., '123'), ensuring data is static and correct."
        )
        return True

    def start_processing(self):
        """Prepares for and launches the background processing thread after getting user input."""
        print("\n--- Settings ---")
        print(f"  Input file: {self.input_file}")
        print(f"  Output directory: {self.output_directory}")
        print(f"  Chunk size: {self.chunk_size}")
        print(f"  Header rows: {self.heading_rows}")
        print(f"  Preserve Formulas: {'Yes' if self.preserve_formulas else 'No'}")
        print("------------------\n")

        try:
            wb_check = openpyxl.load_workbook(self.input_file, read_only=True)
            total_rows_check = wb_check.active.max_row
            wb_check.close()
            data_rows_check = total_rows_check - self.heading_rows
            num_chunks_check = (data_rows_check + self.chunk_size - 1) // self.chunk_size if data_rows_check > 0 else 0
        except Exception as e:
            messagebox.showerror("Error", f"Could not read input file to determine size: {e}")
            return

        if num_chunks_check == 0:
            messagebox.showinfo("Information", "No data rows to process based on settings.")
            return

        self.progress_queue = queue.Queue()
        self.cancel_event = threading.Event()
        
        self.progress_manager = ProgressManager(self.root, "Splitting File...", num_chunks_check, self.cancel_event)
        
        self.worker_thread = threading.Thread(
            target=split_excel_file_with_formatting,
            args=(
                self.input_file, self.output_directory, self.chunk_size, 
                self.heading_rows, self.preserve_formulas,
                self.progress_queue, self.cancel_event
            )
        )
        self.worker_thread.start()
        self.root.after(100, self.check_queue)

    def check_queue(self):
        """Periodically checks the queue for messages from the worker thread and updates the UI accordingly."""
        try:
            message = self.progress_queue.get(block=False)
            if message['type'] == 'progress':
                self.progress_manager.update(message['step'], message['status'])
            elif message['type'] == 'result':
                self.on_task_finished(message['data'])
                return
        except queue.Empty:
            pass

        if self.worker_thread.is_alive():
            self.root.after(100, self.check_queue)
        else:
            self.on_task_finished({'status': 'error', 'message': 'The worker thread terminated unexpectedly.', 'files_created': 'Unknown'})

    def on_task_finished(self, result):
        """
        Handles the final result from the worker, displays a summary, and cleans up resources.
        
        Args:
            result (dict): A dictionary containing the final status, message, and file count.
        """
        if self.progress_manager:
            self.progress_manager.close()

        print("\n--- Operation Summary ---")
        print(f"Status: {result['status'].title()}")
        print(f"Message: {result['message']}")
        print(f"Total Files Created: {result['files_created']}")
        print("-------------------------\n")
        
        status_map = {
            'error': messagebox.showerror,
            'cancelled': messagebox.showwarning,
            'success': messagebox.showinfo
        }
        
        title = result['status'].title()
        message = f"{result['message']}\n\nTotal files created: {result['files_created']}"
        if result['status'] == 'success':
            message = f"{result['message']}\n\nOutput directory:\n{self.output_directory}"

        status_map.get(result['status'], messagebox.showinfo)(title, message)
        
        self.root.destroy()

    def run(self):
        """Main execution flow of the application."""
        print("SplitXL - Excel File Splitter")
        print("=" * 40)
        if self.get_user_input():
            self.start_processing()
        else:
            print("\nOperation cancelled during setup. Exiting...")
            self.root.destroy()

if __name__ == "__main__":
    # Main entry point for the application.
    # It sets up the Tkinter root window and starts the App controller.
    # A try-except block handles cases where a GUI cannot be started.
    try:
        root = tk.Tk()
        root.withdraw()
        app = App(root)
        root.mainloop()
    except tk.TclError as e:
        print(f"Failed to start GUI application: {e}")
        print("This script requires a graphical desktop environment to run.")