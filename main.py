from bs4 import BeautifulSoup
from html.parser import HTMLParser
from html.entities import name2codepoint
import xlsxwriter




def parcing_html_table(table):
    table_dict = []
    row_index = 0
    for row in table.children:
        if row.name != 'tr':
            continue
        cell_index = 0
        if type(row).__name__ != 'NavigableString':
            table_dict.append([])
            table_dict[row_index] = {}
            table_dict[row_index]['context'] = 'table, tr'
            table_dict[row_index]['cells'] = []
            table_dict[row_index]['attrs'] = row.attrs

            for cell in row.children:
                if type(cell).__name__ != 'NavigableString':
                    table_dict[row_index]['cells'].append([])
                    table_dict[row_index]['cells'][cell_index] = {}
                    table_dict[row_index]['cells'][cell_index]['attrs'] = cell.attrs
                    values_index = 0
                    table_dict[row_index]['cells'][cell_index]['values'] = []
                    for cell_values in cell.contents:
                        cell_text = cell_values.string.replace('\n', '').strip()
                        if len(cell_text) > 0:
                            table_dict[row_index]['cells'][cell_index]['values'].append([])
                            table_dict[row_index]['cells'][cell_index]['values'][values_index] = {}
                            table_dict[row_index]['cells'][cell_index]['values'][values_index]['text'] = cell_text
                            if type(cell_values).__name__ != 'NavigableString':
                                cell_class = cell_values.attrs.get('class', None)
                                if cell_class:
                                    cell_class = cell_class[0] 
                                table_dict[row_index]['cells'][cell_index]['values'][values_index]['context'] = cell_class
                            values_index += 1
                    cell_index += 1
            row_index += 1
    return table_dict

def parcing_html_styles(html):
    styles_dict = {}
    for tag in html:
        if tag.name == 'style':
            styles_name = ''
            styles_atrib_name = ''
            styles_atrib_values = ''
            what_we_read = 'styles_name'
            for tag_style_string in tag.children:
                for symbol in tag_style_string:
                    if symbol == '\n':
                        continue
                    elif what_we_read in ('styles_name', 'styles_atrib_name') and symbol == ' ':
                        continue
                    elif symbol == '{':
                        styles_name = styles_name.strip()
                        styles_dict[styles_name] = {}
                        what_we_read = 'styles_atrib_name'
                    elif symbol == ':':
                        styles_atrib_name = styles_atrib_name.strip()
                        styles_dict[styles_name][styles_atrib_name] = []
                        what_we_read = 'styles_atrib_values'
                    elif symbol == ';':
                        styles_atrib_values = styles_atrib_values.strip()
                        if ',' in styles_atrib_values:
                            styles_atrib_values = styles_atrib_values.split(',')
                        else:
                            styles_atrib_values = styles_atrib_values.split(' ')
                        styles_dict[styles_name][styles_atrib_name].append(styles_atrib_values)
                        styles_atrib_name = ''
                        styles_atrib_values = ''
                        what_we_read = 'styles_atrib_name'
                    elif symbol == '}':
                        styles_name = ''
                        what_we_read = 'styles_name'
                    else:
                        if what_we_read == 'styles_name':
                            styles_name += symbol
                        if what_we_read == 'styles_atrib_name':
                            styles_atrib_name += symbol
                        if what_we_read == 'styles_atrib_values':
                            styles_atrib_values += symbol
    return styles_dict


def table_to_excel(table_dict, cursor_start):
    def add_to_set(adding_cells_set, start_x, start_y, cell_end_x, cell_end_y):
        for i in range(start_x, cell_end_x + 1):
            for j in range(start_y, cell_end_y + 1):
                adding_cells_set.add((i, j))
        return adding_cells_set 
    def get_cell_size(cell):
        rowspan = 1
        colspan = 1
        if cell['attrs'].get('rowspan', None):
            rowspan = int(cell['attrs'].get('rowspan', None))
        if cell['attrs'].get('colspan', None):
            colspan = int(cell['attrs'].get('colspan', None))
        return rowspan, colspan  
    def calc_end_of_cell(cell_start_x, cell_start_y, rowspan, colspan):
        cell_end_x = cell_start_x + rowspan - 1
        cell_end_y = cell_start_y + colspan - 1 
        return cell_end_x, cell_end_y 
    def get_cell_text(cell, rowspan, colspan):
        cell_text = ''
        for i in cell['values']:
            if rowspan > 1:
                cell_text += i['text'] + str('\n')
            else:
                cell_text += i['text']
        return cell_text 
    def write_cell(cell_start_x, cell_start_y, cell_end_x, cell_end_y, cell_text):
        if rowspan + colspan > 2:
            if rowspan > 1:
                cell_format = workbook.add_format()
                cell_format.set_text_wrap()
                cell_format.set_align('left')
                cell_format.set_align('top')
                cell_format.set_border() 
                worksheet.merge_range(cell_start_x, cell_start_y, cell_end_x, cell_end_y, cell_text, cell_format)
            else:
                cell_format_two = workbook.add_format()
                cell_format_two.set_border() 
                worksheet.merge_range(cell_start_x, cell_start_y, cell_end_x, cell_end_y, cell_text, cell_format_two)
        else:
            cell_format_two = workbook.add_format()
            cell_format_two.set_border() 
            worksheet.write(cell_start_x, cell_start_y, cell_text, cell_format_two)
    def get_start_next_cell(cell_start_x, cell_start_y, cell_end_y, adding_cells_set):
        cell_start_y = cell_end_y 
        while (cell_start_x, cell_start_y) in adding_cells_set:
            cell_start_y += 1
        return cell_start_y
    def get_start_next_cell_on_new_row(cell_start_x, cell_start_y):
        cell_start_y = 0
        cell_start_x += 1
        while (cell_start_x, cell_start_y) in adding_cells_set:
            cell_start_y += 1
        return cell_start_x, cell_start_y 

    workbook = xlsxwriter.Workbook('hello.xlsx')
    worksheet = workbook.add_worksheet()
    cell_start_x = int(cursor_start[0])
    cell_start_y = int(cursor_start[1])
    adding_cells_set = set()
    
    for row in table_dict:
        for cell in row['cells']:
            rowspan, colspan = get_cell_size(cell)
            cell_end_x, cell_end_y = calc_end_of_cell(cell_start_x, cell_start_y, rowspan, colspan)
            adding_cells_set = set(add_to_set(adding_cells_set, cell_start_x, cell_start_y, cell_end_x, cell_end_y)) 
            cell_text = get_cell_text(cell, rowspan, colspan)
            write_cell(cell_start_x, cell_start_y, cell_end_x, cell_end_y, cell_text)
            cell_start_y = get_start_next_cell(cell_start_x, cell_start_y, cell_end_y, adding_cells_set)
        cell_start_x, cell_start_y = get_start_next_cell_on_new_row(cell_start_x, cell_start_y)
    workbook.close()
        

def main():
    with open('examples\only table.html', 'r', encoding='utf-8') as html_file_code:
        html_code = html_file_code.read()
        

    soup = BeautifulSoup(html_code, 'html.parser')


    table_dict = list(parcing_html_table(soup.table))
    cursor_start = (0, 0)
    table_to_excel(table_dict, cursor_start)

                

if __name__ == "__main__":
    main()    



