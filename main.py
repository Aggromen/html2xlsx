from bs4 import BeautifulSoup
from html.parser import HTMLParser
from html.entities import name2codepoint
import xlsxwriter


with open('examples\complex_example.html', 'r', encoding='utf-8') as html_file_code:
    html_code = html_file_code.read()
    

soup = BeautifulSoup(html_code, 'html.parser')

def print_html_test(table):
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
                        add_text = cell_values.string.replace('\n', '').strip()
                        if len(add_text) > 0:
                            table_dict[row_index]['cells'][cell_index]['values'].append([])
                            table_dict[row_index]['cells'][cell_index]['values'][values_index] = {}
                            table_dict[row_index]['cells'][cell_index]['values'][values_index]['text'] = add_text
                            if type(cell_values).__name__ != 'NavigableString':
                                cell_class = cell_values.attrs.get('class', None)
                                if cell_class:
                                    cell_class = cell_class[0] 
                                table_dict[row_index]['cells'][cell_index]['values'][values_index]['context'] = cell_class
                            values_index += 1
                    cell_index += 1
            row_index += 1
    return table_dict


def table_to_excel(table_dict, start_cell):
    workbook = xlsxwriter.Workbook('hello.xlsx')
    worksheet = workbook.add_worksheet()
    cur_x = int(start_cell[0])
    cur_y = int(start_cell[1])
    set_x_y = set()
    def add_to_set(set_x_y, start_x, start_y, end_x, end_y):
        for i in range(start_x, end_x + 1):
            for j in range(start_y, end_y + 1):
                set_x_y.add((i, j))
        return set_x_y 
    for row in table_dict:
        for cell in row['cells']:
            rowspan = 1
            colspan = 1
            if cell['attrs'].get('rowspan', None):
                rowspan = int(cell['attrs'].get('rowspan', None))
            if cell['attrs'].get('colspan', None):
                colspan = int(cell['attrs'].get('colspan', None))
            end_x = cur_x + rowspan - 1
            end_y = cur_y + colspan - 1 
            set_x_y = set(add_to_set(set_x_y, cur_x, cur_y, end_x, end_y)) 
            add_text = ''
            for i in cell['values']:
                if rowspan > 1:
                    add_text += i['text'] + str('\n')
                else:
                    add_text += i['text']
            if rowspan + colspan > 2:
                if rowspan > 1:
                    cell_format = workbook.add_format()
                    cell_format.set_text_wrap()
                    cell_format.set_align('left')
                    cell_format.set_align('top')
                    cell_format.set_border() 
                    worksheet.merge_range(cur_x, cur_y, end_x, end_y, add_text, cell_format)
                else:
                    cell_format_two = workbook.add_format()
                    cell_format_two.set_border() 
                    worksheet.merge_range(cur_x, cur_y, end_x, end_y, add_text, cell_format_two)
            else:
                cell_format_two = workbook.add_format()
                cell_format_two.set_border() 
                worksheet.write(cur_x, cur_y, add_text, cell_format_two)
            cur_y = end_y
            while (cur_x, cur_y) in set_x_y:
                cur_y += 1
        cur_y = 0
        cur_x += 1
        while (cur_x, cur_y) in set_x_y:
            cur_y += 1
    workbook.close()
        


table_dict = list(print_html_test(soup.table))
start_cell = (0, 0)
table_to_excel(table_dict, start_cell)



