import os
import logging
from file_handlers import read_excel_file
import traceback

def compare_excel_files(file1, file2):
    try:
        content1, sheets1 = read_excel_file(file1)
        content2, sheets2 = read_excel_file(file2)
        
        if not sheets1 and not sheets2:
            return "Both files are empty or could not be read. See the log for detailed information."
        
        if not sheets1:
            return f"File {os.path.basename(file1)} is empty or could not be read. See the log for detailed information."
        
        if not sheets2:
            return f"File {os.path.basename(file2)} is empty or could not be read. See the log for detailed information."
        
        if sheets1 != sheets2:
            return f"The files have different sheets:\nFile 1: {sheets1}\nFile 2: {sheets2}"
        
        result = []
        
        for sheet_name in sheets1:
            logging.info(f"Compare sheets: {sheet_name}")
            
            cells1 = content1[sheet_name]
            cells2 = content2[sheet_name]
            
            if cells1 == cells2:
                result.append(f"Sheet '{sheet_name}' is identical in both files.")
            else:
                result.append(f"Sheet '{sheet_name}' has differences:")
                
                all_cells = set(cells1.keys()) | set(cells2.keys())
                differences = []
                for cell in all_cells:
                    if cell not in cells1:
                        differences.append(f"Cell {cell} is only present in File 2: {cells2[cell]}")
                    elif cell not in cells2:
                        differences.append(f"Cell {cell} is only present in File 1: {cells1[cell]}")
                    elif cells1[cell] != cells2[cell]:
                        differences.append(f"Cell {cell} differs: File 1: {cells1[cell]}, File 2: {cells2[cell]}")
                
                max_differences = 20
                if len(differences) > max_differences:
                    result.extend(differences[:max_differences])
                    result.append(f"... and {len(differences) - max_differences} additional differences.")
                else:
                    result.extend(differences)
        
        return "\n\n".join(result)
    
    except Exception as e:
        logging.error(f"An error occurred while comparing the files: {str(e)}")
        logging.error(traceback.format_exc())
        return f"An error occurred while comparing the files:\n{str(e)}\n\nSee the log file for more information."