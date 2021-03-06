from sys import exit
import mammoth
import os
import string

from bs4 import BeautifulSoup
import pathlib


import re
import pathlib




def rename_file(fpath, new_name):
    new_name_parts = new_name.split('.')

    new_fpath = fpath.replace(fpath.split('\\')[-1],
                              new_name_parts[0] + '.' +
                              new_name_parts[1])

    fpath = pathlib.Path(fpath)

    while 1:
        if 'count' not in locals():
            count = 0

        try:
            if count == 0:
                fpath.rename(new_fpath)
            else:
                new_fpath = new_fpath.replace(new_fpath.split('\\')[-1],
                                              new_name_parts[0] +
                                              ' (' + str(count) + ').' +
                                              new_name_parts[1])
                fpath.rename(new_fpath)

            break
        except FileExistsError as e:
            count += 1

# read data from a file
def read_data(path):
    with open(path, "r") as file_object:
        return file_object.read()
    return 


# write data to a file, overwrite mode
def write_data(path, data, mode='w'):
    with open(path, mode) as file_object:
        file_object.write(data)


image_count = 0


def convert_image(image):

    global image_count

    image_count += 1
    image_name = 'image_' + str(image_count) + '.jpg'

    if not os.path.exists('images'):
        os.makedirs('images')

    image_path = 'images\\' + image_name
    print(image_count)
    with image.open() as image_bytes:
        write_data(image_path, image_bytes.read(), 'wb')

    return {"src": image_path}


def regex_replace(string, pattern, replace_string):

    matches = re.findall(pattern, string)
    if matches:
        print('replace all {} by "{}"'.format(matches, replace_string))
        # for match in matches:
        #    string = string.replace(match, replace_string)
    else:
        print(pattern, ' not found')

    string = re.sub(pattern, replace_string, string)

    return string


def adjust_html_tables(html):
    # find, iter and process each table
    table_tags = re.findall(r'<table.+?</table>', html)
    for table_tag_origin in table_tags:

        # assign table_tag_origin to another val to keep origin for replacement
        table_tag = table_tag_origin
        # a list containing col widths to calculate the ratio of cols
        col_char_lens = {'whole':[], 'one_line':[]}

        # remove redundant tags
        table_tag = re.sub('<strong>|</strong>', '', table_tag)

        # find, iter and process each row in the table
        tr_tags = re.findall(r'<tr.+?</tr>', table_tag)
        for tr_tag in tr_tags:

            # find, iter and process each cell in a col to find the longest string
            td_tags = re.findall(r'<td.+?</td>', tr_tag)
            for x in range(0,len(td_tags)):

                # current cell html
                td_tag = td_tags[x] 

                # pretify the html to check multiline strings
                soup = BeautifulSoup(td_tag, 'html.parser')
                td_tag = soup.prettify()

                # remove all tags of a cell, remove redundant spaces
                txt = re.sub(r'<.+?>','',td_tag)
                txt = re.sub(r' +\n|\n +','\n',txt)

                # find the longest string
                one_line_strings = txt.split('\n')
                max_len_one_line = 0
                for string in one_line_strings:
                    if len(string) > max_len_one_line:
                        max_len_one_line = len(string)
                        max_len_string = string

                #print('-----------------')
                #print(one_line_strings)
                #print(max_len_string)
                #print('-----------------')

                # use this code to use the whole cell len not 1 line
                max_len_whole_cell = sum((len(x) for x in one_line_strings))

                # get longest string len of one line
                try: # get the longest width of cells in a col
                    if col_char_lens['one_line'][x] < max_len_one_line:
                        col_char_lens['one_line'][x] = max_len_one_line
                except IndexError as e: # if the col width not exists yet
                    col_char_lens['one_line'].append(max_len_one_line)

                # get longest string len of whole cell
                try:
                    if col_char_lens['whole'][x] < max_len_whole_cell:
                        col_char_lens['whole'][x] = max_len_whole_cell
                except IndexError as e: # if the col width not exists yet
                    col_char_lens['whole'].append(max_len_whole_cell)


        print(col_char_lens)

        # number of rows to duplicate cells widths
        row_count = table_tag.count('<tr>')
        # number of cols
        col_count = len(col_char_lens['whole'])

        # prepare col widths html and put in a list
        if col_count > 2:
            cell_widths_html = ['<td>']*(col_count)*row_count
        else:
            if col_char_lens['whole'][0] < 25: # the first col maximun character < 30 => 100
                cell_widths_html = (['<td width="100">'] + 
                                    ['<td>']*(col_count-1)
                                    )*row_count
            else:
                cell_widths_html = (['<td width="350">'] + 
                                    ['<td>']*(col_count-1)
                                    )*row_count

        # use format method to put cell_widths_html in to place
        new_table_tag = table_tag_origin.replace('<td>','{}').format(*cell_widths_html)

        html = html.replace(table_tag_origin, new_table_tag)

    return html




def convert_html(user_input_path):

    with open(user_input_path, "rb") as docx_file:
        result = mammoth.convert_to_html(
            docx_file, convert_image=mammoth.images.img_element(convert_image))
        print('converted docx to html')

    html = result.value

    html = html.replace('<table>', '<blockquote><table ' +
                        'width="700" ' +
                        'border="1" ' +
                        'cellspacing="0" ' +
                        'cellpadding="5">')

    html = html.replace('</table>', '</table></blockquote>')
    print('added <blockquote> for all tables and format tables')

    print('added 20 <br> at the end of document')
    html = html + '<br>' * 50

    # remove spaces before and after marks
    html = regex_replace(html, r'\s+<', '<')
    html = regex_replace(html, r'>+\s', '>')


    
    # check if the docx file contains any IDs or Hyperlinks
    if '<a href=' in html or '<a id=' in html:
        print('Found IDs or Hyperlinks in your docx file ' + 
              'Please remove if you want to use automatic putting hyperlinks function')
    else:
        # find bold strings to add IDs
        bold_strings = re.findall(r'<strong>.+?</strong>', html)
        for bold_string in bold_strings:

            # skip bold texts in tables
            if '<td><p>' + bold_string + '</p></td>' in html:
                continue

            text = regex_replace(bold_string, r'<strong>|</strong>', '')
            print('creating id and hyperlinks for: >> ' + text)

            # add id
            new_bold_string = '<a id="' + text + '"></a>' + bold_string
            html = html.replace(bold_string, new_bold_string)
            # add hyperlinks
            html = html.replace('<p>' + text + '</p>', '<p><a href="#{text}">{text}</a></p>'.format(text=text))

 

    # adjust tables
    html = adjust_html_tables(html)



    # add a break before bold text
    html = html.replace(r'<strong>', r'<strong><br>')


    # find, iter and process each cell to delete redandant <br>
    td_tags = re.findall(r'<td.+?</td>', html)
    for td_tag in td_tags:
        new_td_tag = td_tag.replace(r'<strong><br>',
                                    r'<strong>')
        html = html.replace(td_tag, new_td_tag)
    # remove redundant <br> in header
    html = html.replace(r'<h1><strong><br>', r'<h1><strong>')




    soup = BeautifulSoup(html, 'html.parser')
    html = soup.prettify()

    return html


def get_file_paths(*folder_paths):
    '''
    get all files from a folder, 
    also check subfolders 
    '''

    # create a list to contain file paths
    file_paths = []
    while 1:
        '''
        this while loop will break only when no unchecked folders left
        '''
        temp_folder_paths = []  # this list collects folder paths each loop
        for folder_path in folder_paths:  # iter all folder paths
            for path in folder_path.iterdir():  # iter all paths in each folder path
                if path.is_file():  # if path is file
                    file_paths.append(str(path))  # collect
                else:  # folder
                    temp_folder_paths.append(path)  # collect

        # assign the folder paths collect from this turn to folder_paths to run the next turn
        folder_paths = temp_folder_paths

        # if there is no folder paths left => done
        if folder_paths == []:
            break
    return file_paths


def list_filter(data, *, mode='path_contains', string=''):
    if mode == 'path_contains':
        strings = string.split('|')
        result = []
        for string in strings:
            result.extend(list(filter((lambda x: string in x), data)))

    if mode == 'html_contains':
        result = list(
            filter((lambda x: ('.html' in x) and (string in read_data(x))), data))

    if mode == 'incorrect_html_paths':
        result = list(filter(verify_html_path, data))

    return list(set(result))



def backup_file(path):
    try:
        backup_path = path.split('\\')[-1].replace('.', ' backup.')
        rename_file(path, backup_path)
        print('found an existing html file, saved as a backup')
    except Exception as e:
        print('no existing html file, no backup saved')



def capitalize_each_word(html_fp):
    html = read_data(html_fp)
    html = html.title()


    upper_words = ['ECM', 'ECM',
                   'PCM', 'PCM',
                   'EVAP', 'EVAP',
                   'VTEC', 'VTEC',
                   'DTC', 'DTC']
    for word in upper_words:
        pattern = re.compile(word, re.IGNORECASE)
        html = pattern.sub(word.upper(), html)
    
    pattern = re.compile('evapor', re.IGNORECASE)
    html = pattern.sub('Evapor', html)

    backup_file(html_fp)

    write_data(html_fp, html)
    print('done\n\n')







def routine_final_convert(path='', input_interface_flag=True):
    try:
        print('''NOTE:
          - this software converts docx file(s) into html file(s) 
          to shorten your time of compiling DTC routine database,
          - you can input a docx file path, html file path or a folder path 
          which contains docx file(s)
          - support drag and drop item
          - column widths are automatically optimized
          - support removing title numbers
          - automatically save a backup of html file if one exists
          - support capitallizing each word (html file)

          V5 Updates:
          - Auto added IDs and hyperlinks
          - Make sure the docx files do not contain any Bookmarks or Hyperlinks
          ''')

        while 1:
            if input_interface_flag==True:
                user_input_path = input('Input docx, html or folder path: ')
            else:
                user_input_path = path
            
            user_input_path = user_input_path.replace('"', '')

            if '.docx' in user_input_path:
                paths = [user_input_path]

            elif '.html' in user_input_path:
                print('html file found')
                print('capitalize each word')
                capitalize_each_word(user_input_path)
                continue

            else:
                user_input_path = pathlib.Path(user_input_path)
                paths = get_file_paths(user_input_path)
                paths = list_filter(paths, mode='path_contains', string='.docx')
                
                

            print('docx file(s) found:')
            for path in paths:
                print(path)
            print('convert docx files to html files')

            
            
            if input_interface_flag==True:
                remove_title_numbers = input('Do you want to remove all title numbers? (y/n): ')
            else:
                remove_title_numbers = 'n'




            for path in paths:
                print('\n\nchecking: ', path)
                html = convert_html(path)

                # remove "-" at beginning of sentences
                html = html.replace(r'  - ', '  ')

                if 'n' in remove_title_numbers:
                    print('kept all title numbers')
                else:
                    
                    html = regex_replace(html, r' \d+\. ', '')
                    print('replaced all title numbers')
                try:
                    html_fpath = path.replace('.docx', '.html')
                    backup_name = html_fpath.split('\\')[-1].replace('.', ' backup.')
                    rename_file(html_fpath, backup_name)
                    print('found an existing html file, saved as a backup')
                except Exception as e:
                    print('no existing html file, no backup saved')

                write_data(path.replace('.docx', '.html'), html)
                print('done converting')

            print('completed all processes!\n\n')


            # break if loop flag is not set
            if input_interface_flag == False: 
                break
    except Exception as e:
        print('Error found! ', e)
        print('Please check your path!\n\n')
        raise e









def put_tables(path):
    #path = r'C:\Users\ThienNguyen\Desktop\Chrysler\2009 Chrysler Sebring L4, 2.4L\1. Diagnostic Routine\DTC\P2004\P2004.docx'
    html = convert_html(path)
    

    html = regex_replace(html, r'<p>\n Yes\n</p>',
                         r'<table width="700" border="1" cellspacing="0" cellpadding="5"><tr><td width="100"><p>Yes</p></td><td>')


    html = regex_replace(html, r'<p>\n No\n</p>',
                         r'</td></tr><tr><td width="100"><p>No</p></td><td>')

    html = regex_replace(html, r'</p><p>(?=\d+\.)',
                         r'</p></td></tr></table><p>')

    html = regex_replace(html, r'<strong>',
                         r'</table><strong>')

    html = regex_replace(html, r'  +</table><strong>',
                         r'<strong>')

    #soup = BeautifulSoup(html, 'html.parser')
    #html = soup.prettify()

    html = regex_replace(html, r' +<br/>\n','')
    html = regex_replace(html, r'Copyright ©','')

    write_data(path.replace('.docx', '_raw.html'), html)
    print('done converting')


def process_put_tables(docx_path=None):

    while 1:
        try:
            if docx_path == None:
                user_input_path = input('Input docx or folder path: ')
            else:
                user_input_path = docx_path
            #user_input_path = r'C:\Users\ThienNguyen\Desktop\Chrysler\2009 Dodge Grand Caravan V6, 3.3L\1. Diagnostic Routine\DTC'
            #user_input_path = r'C:\Users\ThienNguyen\Desktop\P0404.docx'
            user_input_path = user_input_path.replace('"', '')

            if '.docx' in user_input_path:
                paths = [user_input_path]
            else:
                user_input_path = pathlib.Path(user_input_path)
                paths = get_file_paths(user_input_path)
                paths = list_filter(
                    paths, mode='path_contains', string='.docx')

                print('docx file(s) found:')
                for path in paths:
                    print(path)


            for path in paths:
                print('\n\nchecking: ', path)
                put_tables(path)

            print('completed all processes!\n\n')

        except Exception as e:
            print('Error found! ', e)
            print('Please check your path!\n\n',)

        if docx_path != None: break
#process_put_tables()
#routine_final_convert()


#import pypandoc
#output = pypandoc.convert(source=r'C:\Users\ThienNguyen\Desktop\P0430.html',
#                          format='html',
#                          to='docx',
#                          outputfile='output.docx',
#                          extra_args=['-RTS'])



