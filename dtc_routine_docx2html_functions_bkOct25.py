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
    html = html + '<br>' * 20

    html = regex_replace(html, r'\s+<', '<')
    html = regex_replace(html, r'>+\s', '>')

    matches = re.findall(r'<td>.+?\w<', html)
    for match in matches:
        if len(match) <= 30:
            replace_string = match.replace('<td>', '<td width="100">')
            print('replaced    {}    by    {}   '.format(match, replace_string))

            html = html.replace(match, replace_string)

    # add a break before bold text
    html = html.replace(r'<strong>', r'<strong><br>')
    # remove break in unexpected places => keep title only
    html = html.replace(r'<td width="100"><p><strong><br>',
                        r'<td width="100"><p><strong>')
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







def routine_final_convert(loop=True):
    try:
        print('''NOTE:
          - this software converts docx file(s) into html file(s) 
          to shorten your time of compiling DTC routine database,
          - you can input an docx file path or a folder path 
          which contains docx file(s)
          - support drag and drop item
          - column widths are automatically optimized
          - support removing title numbers
          - automatically save a backup of html file if one exists
          - support capitallizing each word (html file)
          ''')

        while 1:
            #user_input_path = input('Input docx, html or folder path: ')
            user_input_path = r'C:\Users\ThienNguyen\Desktop\DTC P0300.docx'
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

            remove_title_numbers = 'y'
            #remove_title_numbers = input('Do you want to remove all title numbers? (y/n): ')

            for path in paths:
                print('\n\nchecking: ', path)
                html = convert_html(path)

                # remove "-" at beginning of sentences
                html = html.replace(r'  - ', '  ')

                if 'y' in remove_title_numbers:
                    html = regex_replace(html, r' \d+\. ', '')
                    print('replaced all title numbers')
                else:
                    print('kept all title numbers')

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
            if loop == False: 
                break
    except Exception as e:
        print('Error found! ', e)
        print('Please check your path!\n\n',)









def put_tables(path):
    #path = r'C:\Users\ThienNguyen\Desktop\Chrysler\2009 Chrysler Sebring L4, 2.4L\1. Diagnostic Routine\DTC\P2004\P2004.docx'
    html = convert_html(path)

    html = regex_replace(html, r'<p>Yes</p>',
                         r'<table width="700" border="1" cellspacing="0" cellpadding="5"><tr><td width="100"><p>Yes</p></td><td>')

    html = regex_replace(html, r'<p>No</p>',
                         r'</td></tr><tr><td width="100"><p>No</p></td><td>')

    html = regex_replace(html, r'</p><p>(?=\d+\.)',
                         r'</p></td></tr></table><p>')

    soup = BeautifulSoup(html, 'html.parser')
    html = soup.prettify()

    html = regex_replace(html, r' +<br/>\n','')
    html = regex_replace(html, r'Copyright ©','')

    write_data(path.replace('.docx', '_raw.html'), html)
    print('done converting')


def process_put_tables():

    while 1:
        try:
            user_input_path = input('Input docx or folder path: ')
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

#process_put_tables()
#routine_final_convert()


#import pypandoc
#output = pypandoc.convert(source=r'C:\Users\ThienNguyen\Desktop\P0430.html',
#                          format='html',
#                          to='docx',
#                          outputfile='output.docx',
#                          extra_args=['-RTS'])



