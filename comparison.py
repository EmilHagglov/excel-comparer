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
        
        summary = {
            'sheets': {'status': 'Passed' if sheets1 == sheets2 else 'Failed',
                       'count1': len(sheets1), 'count2': len(sheets2)},
            'rows': {'status': 'Checking', 'count1': 0, 'count2': 0},
            'columns': {'status': 'Checking', 'count1': 0, 'count2': 0},
            'content': {'status': 'Checking', 'count1': 0, 'count2': 0}
        }

        result = []
        total_rows1 = total_rows2 = 0
        max_columns1 = max_columns2 = 0
        identical_content = True
        
        for sheet_name in sheets1:
            logging.info(f"Compare sheets: {sheet_name}")
            
            rows1 = get_rows(content1[sheet_name])
            rows2 = get_rows(content2[sheet_name])
            
            total_rows1 += len(rows1)
            total_rows2 += len(rows2)
            max_columns1 = max(max_columns1, max(len(row) for row in rows1) if rows1 else 0)
            max_columns2 = max(max_columns2, max(len(row) for row in rows2) if rows2 else 0)
            
            if set(rows1) != set(rows2):
                identical_content = False
                result.append(f"Sheet '{sheet_name}' has differences:")
                missing_in_file2 = set(rows1) - set(rows2)
                missing_in_file1 = set(rows2) - set(rows1)
                
                if missing_in_file2:
                    result.append(f"Rows in File 1 but not in File 2:")
                    result.extend([f"  Row {rows1.index(row) + 1}: {row[0]}" for row in list(missing_in_file2)[:20]])
                    if len(missing_in_file2) > 20:
                        result.append(f"  ... and {len(missing_in_file2) - 20} more rows")
                if missing_in_file1:
                    result.append(f"Rows in File 2 but not in File 1:")
                    result.extend([f"  Row {rows2.index(row) + 1}: {row[0]}" for row in list(missing_in_file1)[:20]])
                    if len(missing_in_file1) > 20:
                        result.append(f"  ... and {len(missing_in_file1) - 20} more rows")
        
        summary['rows']['count1'] = total_rows1
        summary['rows']['count2'] = total_rows2
        summary['rows']['status'] = 'Passed' if total_rows1 == total_rows2 else 'Failed'
        
        summary['columns']['count1'] = max_columns1
        summary['columns']['count2'] = max_columns2
        summary['columns']['status'] = 'Passed' if max_columns1 == max_columns2 else 'Failed'
        
        summary['content']['status'] = 'Passed' if identical_content else 'Failed'
        
        return summary, "\n\n".join(result)
    
    except Exception as e:
        logging.error(f"An error occurred while comparing the files: {str(e)}")
        logging.error(traceback.format_exc())
        return None, f"An error occurred while comparing the files:\n{str(e)}\n\nSee the log file for more information."

def get_rows(sheet_content):
    rows = []
    for cell, value in sheet_content.items():
        row, col = split_cell_address(cell)
        while len(rows) <= row:
            rows.append([])
        while len(rows[row]) <= col:
            rows[row].append(None)
        rows[row][col] = value
    return [tuple(row) for row in rows if any(cell is not None for cell in row)]

def split_cell_address(address):
    row = int(''.join(filter(str.isdigit, address))) - 1
    col = sum((ord(char) - ord('A') + 1) * (26 ** i) for i, char in enumerate(reversed(address.rstrip('0123456789')))) - 1
    return row, col