import logging
import zipfile
import xml.etree.ElementTree as ET
import traceback
import os

def analyze_xlsx(file_path):
    logging.info(f"Analyzing .xlsx-file: {file_path}")
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            file_list = z.namelist()
            logging.info(f"Files in .xlsx-archive: {file_list}")
            
            sheet_info = []
            if 'xl/workbook.xml' in file_list:
                workbook_xml = z.read('xl/workbook.xml')
                root = ET.fromstring(workbook_xml)
                sheets = root.findall('.//{*}sheet')
                sheet_info = [{"name": sheet.get('name'), "id": sheet.get('sheetId')} for sheet in sheets]
                logging.info(f"Found the following page information: {sheet_info}")
            else:
                logging.warning("Could not find workbook.xml in the .xlsx-file")
            
            shared_strings = []
            if 'xl/sharedStrings.xml' in file_list:
                strings_xml = z.read('xl/sharedStrings.xml')
                strings_root = ET.fromstring(strings_xml)
                shared_strings = [s.text for s in strings_root.findall('.//{*}t')]
                logging.info(f"Number of shared strings: {len(shared_strings)}")
            else:
                logging.warning("Could not find sharedStrings.xml in the .xlsx-file")
            
            sheet_content = {}
            for sheet in sheet_info:
                sheet_file = f"xl/worksheets/sheet{sheet['id']}.xml"
                if sheet_file in file_list:
                    sheet_xml = z.read(sheet_file)
                    sheet_root = ET.fromstring(sheet_xml)
                    cells = sheet_root.findall('.//{*}c')
                    sheet_content[sheet['name']] = {}
                    for cell in cells:
                        cell_ref = cell.get('r')
                        cell_type = cell.get('t')
                        value = cell.find('{*}v')
                        if value is not None:
                            if cell_type == 's':
                                # Shared string
                                idx = int(value.text)
                                cell_value = shared_strings[idx] if idx < len(shared_strings) else ''
                            else:
                                cell_value = value.text
                            sheet_content[sheet['name']][cell_ref] = cell_value
                    logging.info(f"Extracted {len(sheet_content[sheet['name']])} cells from page {sheet['name']}")
                else:
                    logging.warning(f"Could not find {sheet_file} in .xlsx-file")
        
        return sheet_info, sheet_content
    except Exception as e:
        logging.error(f"An error occurred while analyzing the .xlsx file: {str(e)}")
        logging.error(traceback.format_exc())
        return None, None

def read_excel_file(file_path):
    logging.info(f"Attempting to read file: {file_path}")
    
    if file_path.lower().endswith('.xlsx'):
        sheet_info, sheet_content = analyze_xlsx(file_path)
        if sheet_info:
            return sheet_content, [sheet['name'] for sheet in sheet_info]

    logging.error(f"Could not read the file {file_path} with any method")
    return None, []