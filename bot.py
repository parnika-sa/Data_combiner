import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import csv
import time
from collections import OrderedDict

CONFIG_FILE = "config.json"

def is_file_open(filepath):
    try:
        os.rename(filepath, filepath)
        return False
    except:
        return True

def save_paths(source, dest):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"source": source, "dest": dest}, f)

def load_paths():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                return data.get("source"), data.get("dest")
        except:
            return None, None
    return None, None

def move_files():
    source_dir, dest_dir = load_paths()
    if not source_dir or not dest_dir:
        messagebox.showwarning("Missing Paths", "Please set folders first using 'Set Folders'.")
        return
    
    if not os.path.exists(source_dir):
        messagebox.showerror("Error", f"Source folder does not exist: {source_dir}")
        return
    
    if not os.path.exists(dest_dir):
        messagebox.showerror("Error", f"Destination folder does not exist: {dest_dir}")
        return
    
    # Only create data folder
    data_folder = os.path.join(dest_dir, "data")
    os.makedirs(data_folder, exist_ok=True)
    
    excel_counter = 1
    moved_count = 0
    
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            full_path = os.path.join(root, file)
            if is_file_open(full_path):
                continue
            
            ext = os.path.splitext(file)[1].lower()
            try:
                # Only move Excel/CSV files to data folder
                if ext in [".xls", ".xlsx", ".csv"]:
                    base_name = os.path.splitext(file)[0]
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    new_name = f"{base_name}_{timestamp}_{excel_counter}{ext}"
                    excel_counter += 1
                    dest_path = os.path.join(data_folder, new_name)
                    shutil.move(full_path, dest_path)
                    moved_count += 1
            except Exception as e:
                continue
    
    messagebox.showinfo("✅ Complete", f"Files moved successfully!\nTotal files moved: {moved_count}")

def detect_delimiter(file_path):
    """Detect the delimiter used in CSV file"""
    delimiters = [',', ';', '\t', '|']
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
        sample = file.read(1024)
        
    delimiter_count = {}
    for delimiter in delimiters:
        delimiter_count[delimiter] = sample.count(delimiter)
    
    return max(delimiter_count, key=delimiter_count.get) if any(delimiter_count.values()) else ','

def merge_csv_files():
    folder_path = filedialog.askdirectory(title="Select Folder with CSV Files to Merge")
    if not folder_path:
        return
    
    # Find all CSV files
    csv_files = []
    for file in os.listdir(folder_path):
        if file.endswith(('.csv', '.CSV')):
            csv_files.append(file)
    
    if not csv_files:
        messagebox.showwarning("No Files", "No CSV files found in selected folder.")
        return
    
    # Create progress window
    progress_window = tk.Toplevel()
    progress_window.title("Merging CSV Files...")
    progress_window.geometry("400x150")
    progress_window.transient()
    progress_window.grab_set()
    
    progress_label = tk.Label(progress_window, text="Processing files...", font=("Arial", 10))
    progress_label.pack(pady=10)
    
    progress_bar = ttk.Progressbar(progress_window, mode='determinate', length=300)
    progress_bar.pack(pady=10)
    progress_bar['maximum'] = len(csv_files)
    
    status_label = tk.Label(progress_window, text="", font=("Arial", 9))
    status_label.pack(pady=5)
    
    progress_window.update()
    
    total_files = len(csv_files)
    success_files = 0
    error_files = []
    total_rows = 0
    
    output_file_path = os.path.join(folder_path, f"MERGED_ALL_DATA_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    
    try:
        header_written = False
        rows_written = 0
        duplicate_count = 0
        seen_rows = set()
        
        with open(output_file_path, "w", newline='', encoding='utf-8') as outfile:
            writer = csv.writer(outfile)
            
            for i, file_name in enumerate(csv_files):
                file_path = os.path.join(folder_path, file_name)
                status_label.config(text=f"Processing: {file_name}")
                progress_window.update()
                
                try:
                    # Detect delimiter
                    delimiter = detect_delimiter(file_path)
                    
                    # Try different encodings
                    content = None
                    for encoding in ['utf-8', 'latin-1', 'cp1252']:
                        try:
                            with open(file_path, "r", encoding=encoding, errors='ignore') as infile:
                                content = infile.read()
                            break
                        except:
                            continue
                    
                    if not content:
                        error_files.append(f"{file_name}: Could not read file")
                        continue
                    
                    # Parse CSV content
                    lines = content.splitlines()
                    if not lines:
                        continue
                    
                    reader = csv.reader(lines, delimiter=delimiter)
                    rows = list(reader)
                    
                    if not rows:
                        continue
                    
                    # Write header only once (from first file)
                    if not header_written and rows:
                        writer.writerow(rows[0])
                        header_written = True
                    
                    # Write data rows (skip header row)
                    for row in rows[1:] if len(rows) > 1 else []:
                        if row:  # Skip empty rows
                            # Create unique key for duplicate detection
                            row_key = tuple(str(cell).strip().lower() for cell in row if str(cell).strip())
                            
                            if row_key and row_key not in seen_rows:
                                seen_rows.add(row_key)
                                writer.writerow(row)
                                rows_written += 1
                            else:
                                duplicate_count += 1
                    
                    total_rows += len(rows) - 1 if len(rows) > 1 else 0
                    success_files += 1
                    
                except Exception as e:
                    error_files.append(f"{file_name}: {str(e)}")
                
                progress_bar['value'] = i + 1
                progress_window.update()
        
        progress_window.destroy()
        
        # Show detailed summary
        summary = f"""
✅ MERGE COMPLETED SUCCESSFULLY!

📊 STATISTICS:
• Total CSV Files Found: {total_files}
• Successfully Processed: {success_files}
• Total Rows in Source: {total_rows:,}
• Unique Rows Written: {rows_written:,}
• Duplicates Removed: {duplicate_count:,}

📁 OUTPUT FILE:
{os.path.basename(output_file_path)}

💡 Data merged exactly as original - no columns added/modified
        """
        
        if error_files:
            summary += f"\n\n❌ ERRORS ({len(error_files)} files):\n"
            for error in error_files[:3]:
                summary += f"• {error}\n"
            if len(error_files) > 3:
                summary += f"• ... and {len(error_files)-3} more errors"
        
        messagebox.showinfo("Merge Complete", summary)
        
        # Ask if user wants to open the output file location
        if messagebox.askyesno("Open Location", "Do you want to open the output file location?"):
            os.startfile(folder_path)
        
    except Exception as e:
        progress_window.destroy()
        messagebox.showerror("Error", f"❌ Failed to merge files!\n\nError: {str(e)}")

def set_paths():
    source = filedialog.askdirectory(title="Select Source Folder (where files are currently)")
    if not source:
        return
    
    dest = filedialog.askdirectory(title="Select Destination Folder (where to move files)")
    if not dest:
        return
    
    save_paths(source, dest)
    messagebox.showinfo("✅ Saved", f"Paths saved successfully!\n\nSource: {os.path.basename(source)}\nDestination: {os.path.basename(dest)}")

def show_current_paths():
    source, dest = load_paths()
    if source and dest:
        msg = f"📁 CURRENT CONFIGURATION:\n\nSource Folder:\n{source}\n\nDestination Folder:\n{dest}"
        messagebox.showinfo("Current Paths", msg)
    else:
        messagebox.showwarning("No Configuration", "No folder paths configured yet.\n\nPlease use 'Set Folders' first.")

def clear_paths():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)
    messagebox.showinfo("Cleared", "All saved paths cleared successfully!")

# Main GUI
def create_main_gui():
    root = tk.Tk()
    root.title("File Manager & Advanced CSV Combiner v2.0")
    root.geometry("500x400")
    root.resizable(False, False)
    root.configure(bg='#f0f0f0')
    
    # Title
    title_label = tk.Label(root, text="🔧 File Manager & CSV Combiner v2.0", 
                          font=("Arial", 18, "bold"), bg='#f0f0f0', fg='#333')
    title_label.pack(pady=15)
    
    # File Mover Section
    mover_frame = tk.LabelFrame(root, text="📁 File Mover", font=("Arial", 12, "bold"), 
                               padx=15, pady=10, bg='#f0f0f0', fg='#333')
    mover_frame.pack(fill="x", padx=20, pady=10)
    
    tk.Button(mover_frame, text="🛠️ Set Source & Destination Folders", 
              command=set_paths, width=38, height=2, font=("Arial", 10),
              bg='#4285f4', fg='white', relief='raised').pack(pady=5)
    
    path_buttons_frame = tk.Frame(mover_frame, bg='#f0f0f0')
    path_buttons_frame.pack(pady=5)
    
    tk.Button(path_buttons_frame, text="📂 Show Paths", command=show_current_paths, 
              width=18, font=("Arial", 9), bg='#34a853', fg='white').pack(side='left', padx=2)
    
    tk.Button(path_buttons_frame, text="🗑️ Clear Paths", command=clear_paths, 
              width=18, font=("Arial", 9), bg='#ea4335', fg='white').pack(side='right', padx=2)
    
    tk.Button(mover_frame, text="▶️ MOVE FILES NOW", 
              command=move_files, bg="#ff6b35", fg="white", width=38, height=2,
              font=("Arial", 11, "bold"), relief='raised').pack(pady=8)
    
    # CSV Combiner Section
    combiner_frame = tk.LabelFrame(root, text="🔄 Advanced CSV Combiner", font=("Arial", 12, "bold"), 
                                  padx=15, pady=10, bg='#f0f0f0', fg='#333')
    combiner_frame.pack(fill="x", padx=20, pady=10)
    
    tk.Button(combiner_frame, text="🚀 MERGE ALL CSV FILES", 
              command=merge_csv_files, bg="#9c27b0", fg="white", width=38, height=2,
              font=("Arial", 11, "bold"), relief='raised').pack(pady=8)
    
    # Features Info
    features_frame = tk.LabelFrame(root, text="✨ Features", font=("Arial", 11, "bold"), 
                                  padx=10, pady=8, bg='#f0f0f0', fg='#333')
    features_frame.pack(fill="x", padx=20, pady=10)
    
    features_text = """• Smart delimiter detection (comma, semicolon, tab, pipe)
• Handles different encodings (UTF-8, Latin-1, CP1252) 
• Merges data exactly as original - no modifications
• Removes duplicate rows automatically
• Progress bar for large operations
• Maintains original column structure
• Detailed statistics and error reporting"""
    
    tk.Label(features_frame, text=features_text, font=("Arial", 9), 
             justify="left", bg='#f0f0f0', fg='#555').pack(pady=5)
    
    # Footer
    footer_label = tk.Label(root, text="Made with ❤️ for efficient data management", 
                           font=("Arial", 8), bg='#f0f0f0', fg='#888')
    footer_label.pack(side='bottom', pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    create_main_gui()