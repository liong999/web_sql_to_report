import base64
import re
import xlsxwriter
from io import BytesIO
from datetime import datetime, timedelta, date
from odoo import models, fields, api, _
from odoo.exceptions import UserError as Warning

class WebReport(models.TransientModel):
    _name = "web.report"
    _description = "Report Utils"

    @api.model
    def _get_default_datetime_plus_7(self): 
        return datetime.now() + timedelta(hours=7)

    name = fields.Char(string='Name')
    report_file = fields.Binary('File', readonly=True)

    def _excel_col_to_index(self, col):
        """
        Convert Excel column reference (A, B, ..., Z, AA, AB, etc.) to zero-based index
        
        :param col: str - Excel column reference (case-insensitive)
        :return: int - Zero-based column index
        """
        col = col.upper()
        index = 0
        for i, char in enumerate(reversed(col)):
            index += (ord(char) - ord('A') + 1) * (26 ** i)
        return index - 1  # Convert to zero-based index
    
    def _excel_index_to_column_name(self, column_index):
        """
        Converts a 1-based column index to its Excel column name.
        e.g., 1 -> A, 2 -> B, 27 -> AA
        """
        result = ""
        while column_index > 0:
            remainder = (column_index - 1) % 26
            result = chr(65 + remainder) + result  # 65 is ASCII for 'A'
            column_index = (column_index - 1) // 26
        return result

    def custom_title(self,text):
        return re.sub(r"(?:(?<=\W)|^)\w(?=\w)", lambda x: x.group(0).upper(), text)
    
    def get_cell_format(self, value):
        cell_format = 'content'
        if isinstance(value, float):
            cell_format = 'content_float'
        elif isinstance(value, int):
            cell_format = 'content_int'
        elif isinstance(value, date):
            cell_format = 'content_date'
        elif isinstance(value, datetime):
            cell_format = 'content_datetime'
        return self.wbf[cell_format]

    wbf = {}

    def generate_report(self
        , report_name 
        , data
        , data_sheet=False
        , data_summary_header=False
        , start_date=False
        , end_date=False
        , header=True
        , header_color = '#FFFACD'
        , capitalize=True
        , numbering=True
        , auto_filter=True
        , freeze_panes=True
        , freeze_panes_column=0
        , bottom_remark=True
        , show_total_footer=True
        , data_custom_footer={}
        ):
        """ Return XLSX Report from the given 'List of Dictionary' data.
            :param report_name: File name of the report before given datetime at the end of it
            :param data: data with 'List of Dictionary' type, the key will be Header and the value will be inserted into each row 
            :param data_sheet: dictionary of sheet name and data, the key will be sheet name and the value will be inserted into each row 
            :param data_summary_header: dictionary of summary header, the key will be cell name and the value will be inserted into the key cell 
            :param start_date: For report title, if this param is filled, title will generated above header
            :param end_date: For report title, start_date must be filled to show end_date. end date will shown after start_date
            :param header: Enable/Disable header (title) option, header text will generated from data key's, underscores '_' will replaced with space and will capitalize unless all words is uppercase 
            :param header_color: Color of the header, the default will be yellow (#FFFF00)
            :param capitalize: Enable/Disable auto capitalize for header text
            :param numbering: Enable/Disable Numbering. Number will generated at first column. The default will be True
            :param auto_filter: Enable/Disable Auto Filter. 'Auto Filter' will filter first row. The default will be True, however.. this only active when params header is true
            :param freeze_panes: Enable/Disable Freeze Pane. Freeze first column. The default will be True, however.. this only active when params header is true
            :param freeze_panes_column: Freeze column from given integer (starting from 0). The default will be 0, however.. this only active when params header & freeze_panes is true
            :param bottom_remark: Enable/Disable remark Give remark at the bottom of the workbook. The remark contains Downloader name & Download date
            :param show_total_footer: Enable/Disable total footer row. If True, adds a summary row at the bottom with sums of numeric columns. Default is False.
            :param data_custom_footer: Dictionary of custom params formulas for footer. The default is an empty dictionary.
            :return: xlsx file of the report generated
        """
        if not data:
            raise Warning("There is no data available.")

        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)       
        workbook = self.add_workbook_format(workbook,header_color)
        wbf = self.wbf

        if not data_sheet:
            data_sheet = {report_name: data}

        filename = report_name.lower()+"_"+ str(self._get_default_datetime_plus_7())+'.xlsx'
        for sheet_name, data in data_sheet.items():
            # Give Report name
            report_name = sheet_name.replace("/"," ")
            worksheet = workbook.add_worksheet(report_name)

            # Initialize params
            column_size = []
            header_len = len(list(data[0].keys()))
            header_row = 0
            number = 1
            row = 0
            col = 0

            # Handle title
            if start_date:
                # Change header row
                header_row = 3

                # Setup title
                worksheet.merge_range(row,col,row,header_len, report_name, wbf['title_doc'])
                time_title = str(start_date)
                if end_date:
                    time_title += ' - '+str(end_date)
                row += 1
                worksheet.merge_range(row,col,row,header_len, time_title, wbf['title_doc'])

                row += 2
            
            # Handle Summary Header
            highest_summary_row = 0
            summary_column_size = {}
            if data_summary_header:
                for cell, value in data_summary_header.items():
                    # Extract row and column from cell reference (e.g., 'A1' -> col='A', row=1)
                    col_str = ''.join(filter(str.isalpha, cell.upper()))
                    row_str = ''.join(filter(str.isdigit, cell))
                    
                    if not col_str or not row_str:
                        raise Warning("Invalid cell format for summary header. Cell format must be 'column and row'. Ex: A9")
                    
                    try:
                        summary_row = int(row_str) - 1  # Convert to zero-based row
                        summary_col = self._excel_col_to_index(col_str)  # Convert to zero-based column
                        cell_format = self.get_cell_format(value)
                        worksheet.write(summary_row, summary_col, value, cell_format)
                    except (ValueError, IndexError) as e:
                        raise Warning(f"Invalid cell reference: {cell}. Error: {str(e)}")
                    
                    # Change column size if content bigger than previous stored size
                    if summary_column_size.get(summary_col,0) < len(str(value)):
                        summary_column_size[summary_col] = (len(str(value)))
                    
                    # Set highest summary row to determine where the first data will be inserted
                    highest_summary_row = max(highest_summary_row, summary_row)
                
                highest_summary_row += 2
                header_row = highest_summary_row

            row = max(row,highest_summary_row)

            # Handle data
            for line in data:
                col = 0
                # Handle header (First data)
                if header and number == 1:
                    
                    # Give column number
                    if numbering:
                        worksheet.write(row, col, "No", wbf['header'])
                        column_size.append(2)
                        col += 1
                    
                    # Loop key for Header/Title
                    for key in line:
                        # Wirte Header Title with key from dictionary
                        formated_header_string = key.replace('_',' ')
                        # IF capitalize params is true and not all word is uppercase, capitalize words
                        if capitalize and not formated_header_string.isupper():
                            formated_header_string = formated_header_string.capitalize()

                        worksheet.write(row, col, formated_header_string, wbf['header'])
                        
                        # Write initial column size
                        column_size.append(len(str(formated_header_string)))
                        col+=1
                    row +=1
                    col = 0

                # Give column number
                if numbering:
                    worksheet.write(row, col, number, wbf['content'])
                    col += 1

                # Write Content
                for key in line:
                    # Define Cell format
                    cell_data = line[key]
                    cell_format = self.get_cell_format(cell_data)
                    worksheet.write(row, col, cell_data if cell_data else '', cell_format)

                    # Change column size if content bigger than previous stored size
                    current_column_index = list(line.keys()).index(key) + int(numbering)
                    if column_size[current_column_index] < len(str(cell_data)):
                        column_size[current_column_index] = (len(str(cell_data)))

                    col+=1
                
                row +=1
                number +=1
            
            # set column width
            for i in range(0, len(column_size)):
                final_column_size = max(column_size[i], summary_column_size.get(i,0)) + 2
                worksheet.set_column(i, i, final_column_size)

            # set auto_filter (only if worksheet have header)
            if header and auto_filter:
                worksheet.autofilter(header_row, 0, row-1, col-1)

            # freeze panes (only if worksheet have header)
            if header and freeze_panes:
                # Handle freeze_panes_column params, and check format 
                to_freeze_column = 0
                if isinstance(freeze_panes_column,int):
                    to_freeze_column = freeze_panes_column

                worksheet.freeze_panes(header_row+1, to_freeze_column)

            if show_total_footer and data and header:
                # Move to the row after the last data row
                col = 0
                
                # If numbering is enabled, skip the first column
                if numbering:
                    worksheet.write(row, col, "Total", wbf['content_total'])
                    col += 1

                # Get the first data row to determine column types
                first_row = next(iter(data))
                
                # Calculate and write totals for each column
                for i, (key, value) in enumerate(first_row.items()):
                    current_col = col + i
                    # Check if the column contains numeric values
                    if any(isinstance(d[key], (int, float)) for d in data):
                        # Check if there is custom footer calculation send by data_custom_footer params
                        operations = data_custom_footer.get(key.upper(),'SUM')
                        # Get column name from index + Write formula
                        col_name = self._excel_index_to_column_name(current_col+1)
                        formula = f'IFERROR({operations}({col_name}{highest_summary_row+2}:{col_name}{row}),0)'
                        # Write the total with total format
                        worksheet.write_formula(row, current_col, formula, wbf['content_total'])
                    else:
                        # For non-numeric columns, leave empty or put a dash
                        worksheet.write_blank(row, current_col, "", wbf['content_total'])
            
            # Set bottom remark
            if bottom_remark:
                worksheet.merge_range('A%s:D%s'%(row+2,row+2), '%s - %s' % (self.sudo().env.user.name, str(self._get_default_datetime_plus_7())) , wbf['footer']) 
            
        workbook.close()
        out=base64.encodebytes(fp.getvalue())
        report = self.sudo().create({
            'report_file' : out,
            'name' : filename,
        })
        
        fp.close()
        return {
            'type': 'ir.actions.act_url',
            "target": "self",
            'url': '/web/content/web.report/%s/report_file/%s?download=true' % (report.id, filename)
        }  
    
    def add_workbook_format(self, workbook, header_color):
        
        self.wbf['title_doc'] = workbook.add_format({'bold': 1,'align': 'left'})
        self.wbf['title_doc'].set_font_size(12)

        self.wbf['footer'] = workbook.add_format({'align':'left'})

        self.wbf['header'] = workbook.add_format({'bg_color':header_color,'bold': 1,'align': 'center','font_color': '#000000'})
        self.wbf['header'].set_top(2)
        self.wbf['header'].set_bottom()
        self.wbf['header'].set_left()
        self.wbf['header'].set_right()
        self.wbf['header'].set_font_size(11)
        self.wbf['header'].set_align('vcenter')

        self.wbf['content'] = workbook.add_format({'align': 'left','font_color': '#000000'})
        self.wbf['content'].set_left()
        self.wbf['content'].set_right()
        self.wbf['content'].set_top()
        self.wbf['content'].set_bottom()
        self.wbf['content'].set_font_size(10)                

        self.wbf['content_float'] = workbook.add_format({'align': 'right','num_format': '#,##0.00'})
        self.wbf['content_float'].set_right() 
        self.wbf['content_float'].set_left()
        self.wbf['content_float'].set_top()
        self.wbf['content_float'].set_bottom()
        self.wbf['content_float'].set_font_size(10)                
        
        self.wbf['content_total'] = workbook.add_format({'bg_color': header_color, 'bold': 1, 'align': 'right', 'num_format': '#,##0.00'})
        self.wbf['content_total'].set_font_color('#000000')
        self.wbf['content_total'].set_font_size(10)                
        self.wbf['content_total'].set_right() 
        self.wbf['content_total'].set_left()
        self.wbf['content_total'].set_top()
        self.wbf['content_total'].set_bottom()
        
        self.wbf['content_int'] = workbook.add_format({'align': 'right','num_format': '#,##0'})
        self.wbf['content_int'].set_right() 
        self.wbf['content_int'].set_left()
        self.wbf['content_int'].set_top()
        self.wbf['content_int'].set_bottom()
        self.wbf['content_int'].set_font_size(10)
                
        self.wbf['content_date'] = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        self.wbf['content_date'].set_left()
        self.wbf['content_date'].set_right() 
        self.wbf['content_date'].set_top()
        self.wbf['content_date'].set_bottom()
        self.wbf['content_date'].set_font_size(10)                
                
        self.wbf['content_date_time'] = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
        self.wbf['content_date_time'].set_left()
        self.wbf['content_date_time'].set_right() 
        self.wbf['content_date_time'].set_top()
        self.wbf['content_date_time'].set_bottom()
        self.wbf['content_date_time'].set_font_size(10)                
        
        return workbook   
    