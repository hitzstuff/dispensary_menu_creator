import os
import shutil
from datetime import date
import time
import pathlib
import webbrowser
from json import (load as json_load, dump as json_dump)
import openpyxl
import requests
import bs4
import numpy as np
import pandas as pd
import PySimpleGUI as sg
from __init__ import VERSION

# URL to the project page for this program
GITHUB_LINK = 'https://github.com/hitzstuff/dispensary_menu_creator'
# File path of the folder that this program resides in
MAIN_DIRECTORY = str(pathlib.Path( __file__ ).parent.absolute())
CATEGORIES_FILE = pathlib.PurePath(MAIN_DIRECTORY, 'config_files', 'categories.cfg')
# Image and icon file paths
PROGRAM_ICON = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'menu_creator.ico')
ICON_FOLDER = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'folder.png')
ICON_DISCOUNTED_PRODUCTS = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'discounted_products.png')
ICON_PRODUCT_CATEGORIES = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'product_categories.png')
ICON_MENU_ASSIGNMENTS = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'menu_assignments.png')
ICON_MENU_TEMPLATE = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'menu_template.png')
ICON_HELP = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'help.png')
ICON_ABOUT = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'about.png')
ICON_DOWNLOAD = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'download.png')
ICON_GITHUB = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'github.png')
ICON_CONTACT = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'contact.png')
ICON_LINK = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'external_link.png')
ICON_ICONS8 = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'ui', 'icons8.png')
MENU_LOGO = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'menu', 'menu_logo.png')
PROGRAM_LOGO = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'dispensary_logo.png')
# PySimpleGUI theme
THEME = 'SystemDefault1'
# Colors
COL_1_BACKGROUND_COLOR = '#F5F5F5'
COL_2_BACKGROUND_COLOR = '#FFF'
TEXT_LINK_COLOR = '#009678'
TEXT_LINK_COLOR_HOVER = '#00BC96'
BUTTON_COLOR = '#003B48 on #30C095'
BUTTON_COLOR_HOVER = '#003B48 on #59DFAB'
EXIT_COLOR = '#62074A on #E15878'
EXIT_COLOR_HOVER = '#62074A on #F07B8B'
# Dimensions
width, height = sg.Window.get_screen_size()
if height == 1440:
    WINDOW_HEIGHT = 450
    WINDOW_HEIGHT_UPDATE = 500
if height == 1080:
    WINDOW_HEIGHT = 400
    WINDOW_HEIGHT_UPDATE = 450
# Version control/update notification
request = requests.get(GITHUB_LINK, timeout=5)
parse = bs4.BeautifulSoup(request.text, 'html.parser')
parse_part = parse.select('div#readme p')
NEWEST_VERSION = str(parse_part[0])
NEWEST_VERSION = NEWEST_VERSION.split()[-1][:-4]
DOWNLOAD_LINK = str(parse_part[1])
DOWNLOAD_LINK = ((DOWNLOAD_LINK.split()[-1][:-4]).split('>')[1]).split('<')[0]
v1_stop = VERSION.find('-')
v2_stop = NEWEST_VERSION.find('-')
VERSION_1 = VERSION[0:v1_stop]
VERSION_2 = NEWEST_VERSION[0:v2_stop]
if VERSION_2 > VERSION_1:
    AVAILABLE_UPDATE = True
else:
    AVAILABLE_UPDATE = False
MENU_TEMPLATE = pathlib.PurePath(MAIN_DIRECTORY,
                                 'config_files',
                                 'menu_template',
                                 'menu_template.xlsx')
# Pre-configured text for about_window()
ABOUT = (
    f'Current Version:\t{VERSION}\n' +
    f'Newest Version:\t{NEWEST_VERSION}\n\n'
    +
    'Developed by:\tAaron Hitzeman\n' +
    '\t\taaron.hitzeman@gmail.com\n\n\n'
    +
    'User interface icons were obtained from https://icons8.com.\n\n'
    +
    'Please visit the GitHub page for more information.'
)
# Default mapping file, in case a new one needs to be created
DEFAULT_MAPPING = {
    'MMJ Product': '',
    'Unit Price': '',
    'Brand': '',
    'Product Category': '',
    'First Row Number': '',
    'Last Row Number': '',
    'THC Column': '',
    'Type Column': '',
    'Product Column': '',
    }

def mapping_file(page, menu_position):
    '''Returns the mapping file for a specified page and menu position'''
    file = pathlib.PurePath(
            MAIN_DIRECTORY,
            'config_files',
            'worksheet_cell_mapping',
            f'{page}_{menu_position}_map.cfg'
            )
    return file

def save_mapping(page, menu_position, mapping, values):
    '''Saves the cell mapping values for a specified page and menu position'''
    file = mapping_file(page, menu_position)
    if values:
        # Update the window with values read from the mapping file
        i = 0
        for _ in mapping:
            key = _
            value = values[i]
            mapping[key] = value
            i += 1
    # Opens the mappings file and overwrites it with the new values
    with open(file, 'w', encoding='UTF-8') as file:
        json_dump(mapping, file)
    return None

def load_mapping(page, menu_position):
    '''Loads the cell mapping values for a specified page and menu position'''
    try:
        file = mapping_file(page, menu_position)
        with open(file, 'r', encoding='UTF-8') as file:
            mapping = json_load(file)
    except FileNotFoundError:
        mapping = DEFAULT_MAPPING
        save_mapping(page, menu_position, mapping, None)
    return mapping

def unassign_menu(page, menu):
    '''Given a page number and menu letter, replaces the "MMJ Product" value with a blank value'''
    mapping = load_mapping(page, menu)
    mapping['MMJ Product'] = ''
    save_mapping(page, menu, mapping, None)
    return None

def create_menu_file():
    '''Creates a new menu from the template file'''
    today = date.today()
    month = today.strftime('%m')
    day = today.strftime('%d')
    year = today.strftime('%Y')
    name = f'Menu {month}-{day}-{year}.xlsx'
    file_path = pathlib.PurePath(MAIN_DIRECTORY, 'saved_menus', name)
    shutil.copy(MENU_TEMPLATE, file_path)
    menu_file = openpyxl.load_workbook(file_path)
    return menu_file, file_path

def create_window(layout, background_color='#FFF'):
    '''Creates a PySimpleGUI window'''
    window = sg.Window(
        '',
        layout,
        text_justification = 'left',
        font = ('Open Sans', 13),
        background_color = background_color,
        no_titlebar = True,
        finalize = True
        )
    return window

def find_alias(category):
    '''Given a product category, returns its alias'''
    category_dict = load_categories()
    try:
        alias = category_dict[category][0]
    except KeyError:
        alias = category
    return alias

def save_categories(category_dict):
    '''Saves all category information to the categories.cfg file'''
    file = CATEGORIES_FILE
    with open(file, 'w', encoding='UTF-8') as file:
        json_dump(category_dict, file)
    return None

def save_alias(category, alias):
    '''Given a category name and its alias, saves them to categories.cfg'''
    category_dict = load_categories()
    category_dict[category][0] = alias
    save_categories(category_dict)

def category_list():
    '''Collects category names from the worksheet cell mapping files'''
    cat_list = []
    for _ in [1, 2, 3, 4, 5, 6]:
        page = _
        for _ in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            menu = _
            mapping = load_mapping(page, menu)
            category = mapping['MMJ Product']
            if category != '':
                cat_list.append(category)
    return cat_list

def load_categories():
    '''Loads the categories from categories.cfg'''
    try:
        file = CATEGORIES_FILE
        with open(file, 'r', encoding='UTF-8') as file:
            category_dict = json_load(file)
    except FileNotFoundError as error:
        sg.popup(f'exception {error}\n\nNo categories file found.',
                 title='',
                 font = ('Open Sans', 13))
    return category_dict

def load_discounts():
    '''Loads the discount values from the categories.cfg file'''
    category_dict = load_categories()
    discounts = []
    for _ in category_list():
        discounts.append(category_dict[_][1])
    return discounts

def save_discounts(discount_values, overall_discount):
    '''Saves the discount values to their respective categories in the categories.cfg file'''
    category_dict = load_categories()
    cat_list = category_list()
    for i, _ in enumerate(cat_list):
        if _ != '':
            if discount_values[i] == '':
                category_dict[_][1] = overall_discount
            else:
                category_dict[_][1] = discount_values[i]
    file = CATEGORIES_FILE
    # Opens the discount file and overwrites it with the new value
    with open(file, 'w', encoding='UTF-8') as file:
        json_dump(category_dict, file)
    return None

def text_label(text, width, bold=None, style=1):
    '''Returns a text label'''
    if bold is None:
        font_type = ('Open Sans', 12)
    else:
        font_type = ('Open Sans', 12, 'bold')
    if style == 1:
        label = sg.Text(text+': ',
                        justification='l',
                        size =  (width, 1),
                        background_color = '#FFF',
                        text_color = '#1B2D45',
                        pad = ((5, 0), 2),
                        font = font_type
                        )
    if style == 2:
        label = sg.Text(text,
                        justification='l',
                        size =  (width, 1),
                        background_color = '#FFF',
                        text_color = '#1B2D45',
                        pad = ((5, 0), 2),
                        font = font_type
                        )
    return label

def page_swap(page_one, page_two):
    '''Renames config files, swapping the two given page numbers'''
    for menu in ['A', 'B', 'C', 'D', 'E','F', 'G', 'H', 'I']:
        pathlib.PurePath(MAIN_DIRECTORY,
                         'config_files',
                         'worksheet_cell_mapping',
                         f'{page_one}_{menu}_map.cfg'
                         )
        file_one = pathlib.PurePath(MAIN_DIRECTORY,
                                    'config_files',
                                    'worksheet_cell_mapping',
                                    f'{page_one}_{menu}_map.cfg'
                                    )
        file_two = pathlib.PurePath(MAIN_DIRECTORY,
                                    'config_files',
                                    'worksheet_cell_mapping',
                                    f'{page_two}_{menu}_map.cfg'
                                    )
        temp_file = pathlib.PurePath(MAIN_DIRECTORY,
                                     'config_files',
                                     'worksheet_cell_mapping',
                                     f'{page_two}_{menu}_map.txt'
                                     )
        os.rename(file_one, temp_file)
        os.rename(file_two, file_one)
        os.rename(temp_file, file_two)
    workbook = openpyxl.load_workbook(filename = MENU_TEMPLATE)
    workbook.active = workbook[f'page_{max([page_one, page_two])}']
    workbook.move_sheet(workbook.active, offset = -1)
    workbook[f'page_{page_one}'].title = f'page_{page_two}x'
    workbook[f'page_{page_two}'].title = f'page_{page_one}'
    workbook[f'page_{page_two}x'].title = f'page_{page_two}'

def menu_swap(menu_one, menu_two):
    '''Given two valid menus, will swap the assigned categories'''
    map_one = load_mapping(int(menu_one[0]), menu_one[1])
    map_two = load_mapping(int(menu_two[0]), menu_two[1])
    category_one = map_one['MMJ Product']
    category_two = map_two['MMJ Product']
    map_one['MMJ Product'] = category_two
    map_two['MMJ Product'] = category_one
    save_mapping(int(menu_one[0]), menu_one[1], map_one, None)
    save_mapping(int(menu_two[0]), menu_two[1], map_two, None)

def menu_locations(cat_type=None):
    '''Returns a list of menus and their product categories'''
    menu_dict = {}
    for _ in [1, 2, 3, 4, 5, 6]:
        page = _
        for _ in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            menu = _
            mapping = load_mapping(page, menu)
            if cat_type != 'alias':
                category = mapping['MMJ Product']
            else:
                alias = find_alias(mapping['MMJ Product'])
                category = alias
            menu_dict[category] = f'{page}{menu}'
    return menu_dict

def table_categories():
    '''Populates a list of categories to act as values for a PySimpleGUI table'''
    category_names = load_categories()
    categories_list = []
    for _ in category_names.items():
        category = find_alias(_[0])
        categories_list.append(category)
    locations = menu_locations()
    categories = []
    for _ in locations.items():
        page = int(_[1][0])
        menu = _[1][1]
        mapping = load_mapping(page, menu)
        name = mapping['MMJ Product']
        alias = find_alias(name)
        if name != '':
            categories.append(f'"({page}) {menu}\t{alias}"')
    categories = np.array(categories)
    np.reshape(categories, (len(categories), 1))
    return categories, categories_list

def move_menu_layout(move_type):
    '''Creates the layout for moving menus'''
    sg.theme(THEME)
    if move_type == 'page':
        method = 'Page'
        syntax_string = '1, 2, 3, 4, 5, 6'
    if move_type == 'menu':
        method = 'Menu'
        syntax_string = '1A, 1B, 1C, 2A, 2B'
    layout = [
        [text_label(
            f'{method} #1', 15, 'bold'),
            sg.Input(
                key = '-METHOD_1-',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                justification = 'c',
                pad = ((25, 10), 2))
            ],
        [text_label(
            f'{method} #2', 15, 'bold'),
            sg.Input(
                key = '-METHOD_2-',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                justification = 'c',
                pad = ((25, 10), 2))
            ],
        [sg.Text(f'Proper syntax:\n{syntax_string}',
                 font = ('Open Sans', 11, 'bold'),
                 text_color = '#0D70E8',
                 pad = (0, 10),
                 border_width = 0,
                 background_color = '#FFF'
                 )],
        [
        sg.Button(
            'Swap',
            size = (10, 1),
            enable_events = True,
            key = '-SWAP_MENUS-',
            button_color = '#003B48 on #30C095',
            font = ('Open Sans', 11, 'bold'),
            pad = (5, (25, 15))
            ),
        sg.Button(
                'Exit',
                size = (10, 1),
                enable_events = True,
                key = '-EXIT-',
                button_color = '#62074A on #E15878',
                font = ('Open Sans', 11, 'bold'),
                pad = (5, (25, 15))
                )
            ]
        ]
    return layout

def move_menu(move_type):
    '''Handle events for the cell mapping window'''
    window = None
    while True:
        try:
            if window is None:
                layout = move_menu_layout(move_type)
                window = create_window(layout)
                window['-SWAP_MENUS-'].bind('<Enter>', 'ENTER')
                window['-SWAP_MENUS-'].bind('<Leave>', 'EXIT')
                window['-EXIT-'].bind('<Enter>', 'ENTER')
                window['-EXIT-'].bind('<Leave>', 'EXIT')
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '-EXIT-':
                window.close()
                break
            if event == '-SWAP_MENUS-ENTER':
                window['-SWAP_MENUS-'].update(button_color='#003B48 on #59DFAB')
                window.set_cursor('hand2')
            if event == '-SWAP_MENUS-EXIT':
                window['-SWAP_MENUS-'].update(button_color='#003B48 on #30C095')
                window.set_cursor('arrow')
            if event == '-EXIT-ENTER':
                window['-EXIT-'].update(button_color='#62074A on #F07B8B')
                window.set_cursor('hand2')
            if event == '-EXIT-EXIT':
                window['-EXIT-'].update(button_color='#62074A on #E15878')
                window.set_cursor('arrow')
            if event == '-SWAP_MENUS-':
                window['-SWAP_MENUS-'].update(disabled=True)
                window.set_cursor('wait')
                if len(values['-METHOD_1-']) and len(values['-METHOD_2-']) == 1:
                    page_swap(int(values['-METHOD_1-']), int(values['-METHOD_2-']))
                else:
                    menu_swap(values['-METHOD_1-'], values['-METHOD_2-'])
                window['-SWAP_MENUS-'].update(disabled=False)
                window.set_cursor('hand2')
                window.close()
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 13)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 13)
                )
    window.close()

def unassigned_categories():
    '''Returns a list of unassigned categories'''
    categories = load_categories()
    category_list = []
    alias_list = []
    for _ in categories:
        category_list.append(_)
    for _ in [1, 2, 3, 4, 5, 6]:
        page = _
        for _ in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            menu = _
            mapping = load_mapping(page, menu)
            mapping_category = mapping['MMJ Product']
            if mapping_category in category_list:
                category_list.remove(mapping_category)
    for _ in category_list:
        alias = find_alias(_)
        alias_list.append(alias)
    return alias_list

def range_char(start, stop):
    '''Equivalent to Python's built-in range function, but for letters instead of numbers'''
    converted_characters = (chr(_) for _ in range(ord(start), ord(stop) + 1))
    return converted_characters

def assigned_menu_locations():
    '''Returns two dictionaries: one for unassigned menus, and one for those already assigned'''
    unassigned_menus = {}
    assigned_menus = {}
    for _ in [1, 2, 3, 4, 5, 6]:
        page = _
        for _ in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            menu = _
            mapping = load_mapping(page, menu)
            if mapping['MMJ Product'] == '':
                if page in unassigned_menus:
                    unassigned_menus[page].append(menu)
                else:
                    unassigned_menus[page] = [menu]
            else:
                if page in assigned_menus:
                    assigned_menus[page].append(menu)
                else:
                    assigned_menus[page] = [menu]
    return unassigned_menus, assigned_menus

def find_category_name(alias):
    '''Given a product alias, returns its category name'''
    categories = load_categories()
    category_dict = {}
    for _ in categories:
        key = categories[_][0]
        value = _
        category_dict[key] = value
    return category_dict[alias]

def text_button(text, background_color, text_color, style=1):
    '''Given text, returns an element that looks like a button'''
    if style == 1:
        button = sg.Text(text,
                        key = ('-B-', text),
                        enable_events = True,
                        justification = 'r',
                        background_color = background_color,
                        font = ('Open Sans', 11, 'bold underline'),
                        pad = (5, 0),
                        text_color = text_color)
    if style == 2:
        button = sg.Text(text,
                        key = ('-B-', text),
                        relief = 'raised',
                        enable_events = True,
                        background_color = background_color,
                        font = ('Open Sans', 11, 'bold'),
                        text_color = text_color)
    return button

def bind_button(window, button_text):
    '''Magic code that enables mouseover highlighting to work'''
    _ = button_text
    window[('-B-', _)].bind('<Enter>', 'ENTER')
    window[('-B-', _)].bind('<Leave>', 'EXIT')

def df_clean(data):
    '''Cleans up a DataFrame by removing unnecessary and duplicate data'''
    data.columns = data.columns.str.replace(' ', '_')
    data.columns = data.columns.str.replace('-', '_')
    data.columns = data.columns.str.replace('%', '')
    data.rename(columns=lambda x: x.lower(), inplace=True)
    unlocked_packages = data[data.available >= 1].reset_index(drop=True)
    locked_packages = data[data.lock_code == 'Newly Received'].reset_index(drop=True)
    available_products = [unlocked_packages, locked_packages]
    dataframe = pd.concat(available_products)
    dataframe = dataframe.filter(
        [
        'sku_retail_display_name',
        'sku_name',
        'strain',
        'unit_price',
        'thc',
        'category',
        'available',
        'lock_code'
        ], axis=1).reset_index(drop=True)
    dataframe['thc'] = dataframe['thc'].round(1)
    dataframe.category = dataframe.category.replace('Raw Pre-Roll', 'Pre-Roll')
    dataframe.category = dataframe.category.replace('Vape Cart Distillate', 'Vape Cart')
    dataframe = dataframe.drop_duplicates(subset=['sku_retail_display_name'])
    dataframe = dataframe.reset_index(drop=True)
    return dataframe

def df_fix(cleaned_data):
    '''Fixes labeling issues in poorly-configured inventory databases'''
    # Fix the category names
    brand_list = []
    category_list = []
    for i, _ in enumerate(cleaned_data.sku_name):
        sku = _
        brand = sku.split()[0]
        brand_list.append(brand)
        dash_1 = sku.find(' - ')
        dash_2 = sku[dash_1 + 2:].find(' - ') + dash_1
        category_name = sku[dash_1 + 3:dash_2 + 2].title()
        dash_3 = sku[dash_2 + 2:].find(' - ') + dash_2
        dash_4 = sku[dash_3 + 5:].find(' - ') + dash_3
        category_size = sku[dash_3 + 5:dash_4 + 5].lower()
        category_list.append(f'{brand} {category_size} {category_name}')
    cleaned_data.drop('category', axis=1)
    cleaned_data['category'] = category_list
    cleaned_data['brand'] = brand_list
    # Fix the strain names
    type_list = ['(Hybrid)', '(Indica)', '(Sativa)']
    strain_list = []
    for i, _ in enumerate(cleaned_data.strain):
        strain = _
        display_name = cleaned_data.sku_retail_display_name.iloc[i]
        if str(_) == 'nan':
            display_name = cleaned_data.sku_retail_display_name.iloc[i]
            if display_name.split()[-1] in type_list:
                strain = display_name.split()[-1]
                strain = (strain.replace('(', '')).replace(')', '') + ' Blend'
            elif display_name.split()[-1][-2:] == 'ct':
                category_name = cleaned_data.category.iloc[i].split()[2]
                category_start = str(display_name).find(category_name)
                category_size = cleaned_data.category.iloc[i].split()[1]
                size_start = str(display_name).find(category_size) + len(category_size) + 1
                strain = display_name[size_start:category_start - 1]
            else:
                strain = (display_name.split()[-1].replace('(', '')).replace(')', '')
        elif str(_) == 'THC':
            category_name = cleaned_data.category.iloc[i].split()[2]
            category_start = str(display_name).find(category_name)
            category_size = cleaned_data.category.iloc[i].split()[1]
            size_start = str(display_name).find(category_size) + len(category_size) + 1
            strain = display_name[size_start:category_start - 1]
        else:
            strain = _
        strain_list.append(strain)
    cleaned_data.drop('strain', axis=1)
    cleaned_data['strain'] = strain_list
    # Fix the THC amounts
    thc_list = []
    for i, _ in enumerate(cleaned_data.thc):
        thc = _
        display_name = cleaned_data.sku_retail_display_name.iloc[i]
        if thc < 5:
            if display_name.split()[1][-2:] == 'mg':
                thc = display_name.split()[1]
            else:
                thc = (display_name.split()[2].replace('(', '')).replace(')', '')
        thc_list.append(thc)
    cleaned_data.drop('thc', axis=1)
    cleaned_data['thc'] = thc_list
    category_list = []
    for i, _ in enumerate(cleaned_data.category):
        category = cleaned_data.category.iloc[i]
        data = cleaned_data[cleaned_data.category == _]
        data = list(data.unit_price.unique())
        if len(data) > 1:
            count = cleaned_data.iloc[i].sku_name.split(' - ')[3]
            category = f'{category} {count}'
        category_list.append(category)
    cleaned_data.drop('category', axis=1)
    cleaned_data['category'] = category_list
    cleaned_data = cleaned_data.filter(
        [
        'sku_retail_display_name',
        'brand',
        'strain',
        'unit_price',
        'thc',
        'category',
        ], axis=1).reset_index(drop=True)
    return cleaned_data

def populate_categories(menu):
    '''Populates categories from a new menu and adds them to the current dictionary'''
    categories_list = list(menu.category.unique())
    categories = load_categories()
    for _ in categories_list:
        if _ not in categories:
            categories[_] = [_, 0, '']
    save_categories(categories)

def build_menu(dataframe):
    '''Organizes a DataFrame into a structure suitable for a menu'''
    cleaned_packages = df_clean(dataframe)
    cleaned_packages = df_fix(cleaned_packages)
    dataframe = cleaned_packages
    brands = []
    sizes = []
    word_counts = []
    for i, _ in enumerate(dataframe.sku_retail_display_name.unique()):
        brand = str(_).split()[0].replace("['", "")
        size = str(_).split()[1]
        word_count = len(_)
        brands.append(brand)
        sizes.append(size)
        word_counts.append(word_count)
    dataframe['brand'] = brands
    dataframe['product_size'] = sizes
    dataframe['word_count'] = word_counts
    brand = dataframe['brand']
    product_types = []
    for i, _ in enumerate(dataframe.sku_retail_display_name.unique()):
        par_beg0 = _.find('(') + 1
        par_end0 = _.find(')')
        par_beg1 = _.find('(', par_beg0 + 1) + 1
        par_end1 = _.find(')', par_end0 + 1)
        if '(' in str(_[par_beg1:par_end1]):
            product_types.append(_[par_beg0:par_end0])
        elif '(' not in _:
            sku_parts = _.split()
            for i, part in enumerate(sku_parts):
                if part == 'Blend':
                    list_pos = i - 1
            product_type = sku_parts[list_pos]
            product_types.append(product_type)
        else:
            product_types.append(_[par_beg1:par_end1])
            dataframe['thc'][i] = str(_[par_beg0:par_end0])
    dataframe['product_type'] = product_types
    dataframe = dataframe.filter(
                [
                'sku_retail_display_name',
                'unit_price',
                'brand',
                'size',
                'category',
                'product_type',
                'strain',
                'thc'
                ], axis=1).reset_index(drop=True)
    dataframe = dataframe.query('brand != "EPO"').reset_index(drop=True)
    dataframe.sort_values(by = ['category', 'strain'], inplace=True)
    dataframe.sort_values(by = ['unit_price', 'thc'], ascending=False, inplace=True)
    populate_categories(dataframe)
    return dataframe

def save_product_type(sheet, menu_category, column, first_row, last_row):
    '''Saves the product types to the menu'''
    prod_type = []
    for _ in list(menu_category['product_type']):
        prod_type.append(_)
    row = first_row
    for _ in prod_type:
        cell = f'{column}{row}'
        try:
            sheet[cell].value = _
            if row < last_row:
                row += 1
                cell = f'{column}{row}'
        except AttributeError:
            row += 1
    while row < last_row:
        cell = f'{column}{row}'
        try:
            sheet[cell].value = ''
            row += 1
        except AttributeError:
            row += 12
    return None

def save_product_thc(sheet, menu_category, column, first_row, last_row):
    '''Saves the THC % to the menu'''
    thc = []
    for _ in list(menu_category['thc']):
        thc.append(_)
    row = first_row
    for _ in thc:
        cell = f'{column}{row}'
        try:
            sheet[cell].value = _
            if row < last_row:
                row += 1
                cell = f'{column}{row}'
        except AttributeError:
            row += 1
    while row < last_row:
        cell = f'{column}{row}'
        try:
            sheet[cell].value = ''
            row += 1
        except AttributeError:
            row += 1
    return None

def save_product(sheet, menu_category, column, first_row, last_row):
    '''Saves the product names to the menu'''
    product = []
    for _ in list(menu_category['strain']):
        product.append(_)
    row = first_row
    for _ in product:
        cell = f'{column}{row}'
        try:
            sheet[cell].value = _
            if row < last_row:
                row += 1
                cell = f'{column}{row}'
        except AttributeError:
            row += 1
    while row < last_row:
        cell = f'{column}{row}'
        try:
            sheet[cell].value = ''
            row += 1
        except AttributeError:
            row += 1
    return sheet

def cell_locations(page_number, menu_letter):
    '''Populates the cell locations within a given menu'''
    mapping = load_mapping(page_number, menu_letter)
    menu_unit_price = mapping['Unit Price']
    brand = mapping['Brand']
    menu_category = mapping['Product Category']
    thc_col = mapping['THC Column']
    type_col = mapping['Type Column']
    prod_col = mapping['Product Column']
    first_row = mapping['First Row Number']
    last_row = mapping['Last Row Number']
    category = mapping['MMJ Product']
    alias = find_alias(category)
    menu_cell_map = [menu_unit_price,
                     brand,
                     menu_category,
                     thc_col,
                     type_col,
                     prod_col,
                     first_row,
                     last_row,
                     alias]
    return menu_cell_map

def save_menu(workbook,
              workbook_path,
              full_menu,
              menu_category,
              page_number,
              menu_letter,
              sale_percent=0):
    '''Saves a populated menu to a given location in a worksheet'''
    sheet = workbook[f'page_{page_number}']
    alias = find_alias(menu_category)
    # Collect a list of cell positions for the menu
    menu_cell_map = cell_locations(page_number, menu_letter)
    # Dissect the list into easy-to-read variables
    unit_price_pos = menu_cell_map[0]
    brand_pos = menu_cell_map[1]
    category_pos = menu_cell_map[2]
    thc_col = menu_cell_map[3]
    type_col = menu_cell_map[4]
    prod_col = menu_cell_map[5]
    first_row = int(menu_cell_map[6])
    last_row = int(menu_cell_map[7])
    # Filter the entire menu, selecting products from a specific category
    menu = full_menu[full_menu.category == menu_category]
    # Populate the variables that are consistently singular
    if len(menu) >= 0:
        try:
            discount_cell = f'{prod_col}{first_row - 1}'
            unit_price = float(list(menu['unit_price'])[0])
            if sale_percent == 0:
                discount_msg = ''
                sheet[discount_cell].value = discount_msg
            if sale_percent != 0:
                sale_price = float(list(menu['unit_price'])[0]) * (1 - (sale_percent / 100))
                discount_msg = f'${sale_price:.2f} - {sale_percent}% OFF!'
                sheet[discount_cell].value = discount_msg
            product_brand = list(menu['brand'])[0]
        except ValueError:
            unit_price = ''
            product_brand = ''
            discount_msg = ''
        except IndexError:
            unit_price = ''
            product_brand = ''
            discount_msg = ''
        # Overwrites the menu with new values
        save_product_type(sheet, menu, type_col, first_row, last_row)
        save_product_thc(sheet, menu, thc_col, first_row, last_row)
        save_product(sheet, menu, prod_col, first_row, last_row)
        sheet[unit_price_pos].value = unit_price
        sheet[category_pos].value = alias
        sheet[brand_pos].value = product_brand
        # Save the new values to the workbook
        workbook.save(workbook_path)
        workbook.close()
    return None

def find_discount(category):
    '''Given a product category, returns the current discount percent'''
    categories = load_categories()
    discount = categories[category][1]
    if discount == '':
        discount = 0
    discount = int(discount)
    return discount

def save_all(full_menu, window):
    '''Saves each menu to the Excel workbook'''
    workbook, workbook_path = create_menu_file()
    pages = [1, 2, 3, 4, 5, 6]
    menus = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    for i, _ in enumerate(pages):
        progress_amount = (100 / len(pages)) + (i * 100) / len(pages)
        window['-P_PERCENT-'].update(f'{progress_amount:.1f}%')
        window['-P_BAR-'].update(current_count = progress_amount)
        page = _
        sheet = workbook[f'page_{_}']
        menu_logo = openpyxl.drawing.image.Image(str(MENU_LOGO))
        menu_logo.anchor = 'B1'
        sheet.add_image(menu_logo)
        workbook.save(workbook_path)
        workbook.close()
        for _ in menus:
            menu = _
            mapping = load_mapping(page, menu)
            category = mapping['MMJ Product']
            if category != '':
                discount = find_discount(category)
                save_menu(workbook, workbook_path, full_menu, category, page, menu, discount)
    window['-FILE_NAME-'].update('Menu was successfully created')
    return None

def cell_map_layout():
    '''Layout for worksheet cell mapping'''
    categories, categories_list = table_categories()
    categories_list = unassigned_categories()
    sg.theme(THEME)
    headings = [
    'Page',
    'Menu',
    'Product Category'
    ]
    rows = []
    for _ in categories:
        menu_position = _.split('\t')[0]
        page = menu_position.split()[0][2]
        menu = menu_position.split()[1]
        alias = _.split('\t')[1][:-1]
        row = [page, menu, alias]
        rows.append(row)
    menu_category_column = [
        [sg.Table(values = rows,
                  headings = headings,
                  justification = 'left',
                  key = '-TABLE-',
                  enable_events = True,
                  size = (len(categories), len(categories)),
                  background_color = '#F7F9FC',
                  text_color = '#1B2D45',
                  font = ('Open Sans', 12),
                  selected_row_colors = '#003B48 on #C7F9DC',
                  header_background_color = '',
                  header_text_color = '',
                  pad = 0,
                  max_col_width = 25,
                  select_mode = 'extended',
                  auto_size_columns = True
                  )]]
    menu_mapping_column = [
        [sg.Text('Unassigned Categories:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 pad = (5, 0),
                 border_width = 0,
                 background_color = '#FFF'
                 )],
        [sg.Listbox(values = categories_list,
                 key = '-MMJ_PRODUCT-',
                 enable_events = True,
                 font = ('Open Sans', 12),
                 text_color = '#1A2138',
                 background_color = '#F7F9FC',
                 highlight_background_color = '#C7F9DC',
                 highlight_text_color = '#003B48',
                 pad = (5, 0),
                 size = (33, 6)
                 )],
        [sg.Text('', background_color = '#FFF')],
        [sg.Text('Page Number:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 pad = (5, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 )],
        [sg.Button(f'{num}',
                   key = f'{num}',
                   enable_events = True,
                   size = (3, 1),
                   font = ('Open Sans', 11, 'bold'),
                   button_color = '#003B48 on #C7F9DC',
                   disabled_button_color = '#003B48 on #30C095'
                   ) for num in range(1, 7)
         ],
        [sg.Text('', background_color = '#FFF')],
        [sg.Text('Menu Position:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 pad = (5, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 )],
        [sg.Button(f'{letter}',
                   key = f'{letter}',
                   enable_events = True,
                   size = (3, 1),
                   font = ('Open Sans', 11, 'bold'),
                   button_color = '#003B48 on #C7F9DC',
                   disabled_button_color = '#003B48 on #30C095'
                   ) for letter in range_char('A', 'C')
         ],
        [sg.Button(f'{letter}',
                   key = f'{letter}',
                   enable_events = True,
                   size = (3, 1),
                   font = ('Open Sans', 11, 'bold'),
                   button_color = '#003B48 on #C7F9DC',
                   disabled_button_color = '#003B48 on #30C095'
                   ) for letter in range_char('D', 'F')
         ],
        [sg.Button(f'{letter}',
                   key = f'{letter}',
                   enable_events = True,
                   size = (3, 1),
                   font = ('Open Sans', 11, 'bold'),
                   button_color = '#003B48 on #C7F9DC',
                   disabled_button_color = '#003B48 on #30C095'
                   ) for letter in range_char('G', 'I')
         ],
        [sg.Text('', background_color = '#FFF')],
        [sg.Text('Location on Worksheet:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 pad = (5, 5),
                 border_width = 0,
                 background_color = '#FFF'
                 )],
        [text_label(
            'Unit Price', 16),
            sg.Input(
                key = '-UNIT_PRICE-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                pad = ((0, 15), 2))
            ],
        [text_label(
        'Brand', 16),
        sg.Input(
            key = '-BRAND-',
            justification = 'c',
            size = (4, 1),
            enable_events = True,
            background_color = '#F7F9FC',
            pad = ((0, 15), 2))
        ],
        [text_label(
            'Product Category', 16),
            sg.Input(
                key = '-CATEGORY-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'First Row Number', 16),
            sg.Input(
                key = '-ROW_START-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'Last Row Number', 16),
            sg.Input(
                key = '-ROW_END-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'THC Column', 16),
            sg.Input(
                key = '-THC_COL-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'Type Column', 16),
            sg.Input(
                key = '-TYPE_COL-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'Product Column', 16),
            sg.Input(
                key = '-PRODUCT_COL-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#F7F9FC',
                pad = ((0, 15), 2))
            ]
        ]
    title_column_l = [
        [sg.Text('',
                 enable_events = True,
                 key = '-TITLE_L-',
                 background_color = '#FFF',
                 text_color = '#0D70E8',
                 justification = 'left',
                 pad = ((0, 5), (5, 0)),
                 font = ('Open Sans', 13, 'bold')
                 )]
    ]
    title_column_r = [
        [sg.Text('',
                 enable_events = True,
                 key = '-TITLE_R-',
                 background_color = '#FFF',
                 text_color = '#0D70E8',
                 justification = 'left',
                 pad = (5, (5, 0)),
                 font = ('Open Sans', 13, 'bold')
                 )]
    ]
    layout = [
        [sg.Column(title_column_l,
                   background_color = '#FFF',
                   size = (170, 90),
                   pad = ((10, 20), (5, 0))
                   ),
         sg.Column(title_column_r,
                   background_color = '#FFF',
                   justification = 'left',
                   size = (410, 90),
                   pad = ((10, 10), (5, 0))
                   )
         ],
          [
            sg.Button('Page Swap',
                    size = (10, 1),
                    enable_events = True,
                    button_color = '#003B48 on #30C095',
                    font = ('Open Sans', 11, 'bold'),
                    key = '-PAGE_SWAP-',
                    pad = ((5, 5), (0, 15))
                    ),
            sg.Button('Menu Swap',
                    size = (10, 1),
                    enable_events = True,
                    button_color = '#003B48 on #30C095',
                    font = ('Open Sans', 11, 'bold'),
                    key = '-MENU_SWAP-',
                    pad = ((5, 5), (0, 15))
                    ),
            sg.Button('Unassign Menu',
                    size = (14, 1),
                    enable_events = True,
                    key = '-UNASSIGN_MENU-',
                    button_color = '#62074A on #E15878',
                    font = ('Open Sans', 11, 'bold'),
                    pad = ((162, 0), (0, 15))
            )
        ],
        [sg.Column(menu_category_column,
                   pad = ((0, 10), (10, 0)),
                   vertical_alignment = 'top'),
         sg.Column(menu_mapping_column,
                   pad = (10, (10, 0)),
                   vertical_alignment = 'top',
                   background_color = '#FFF'),
         ],
        [sg.Button(
                'Save',
                size = (7, 1),
                enable_events = True,
                key = '-SAVE-',
                button_color = '#003B48 on #30C095',
                font = ('Open Sans', 11, 'bold'),
                pad = ((625, 5), (15, 0))
                )
            ],
        [sg.Button(
            'Exit',
            size = (7, 1),
            enable_events = True,
            key = '-EXIT-',
            button_color = '#62074A on #E15878',
            font = ('Open Sans', 11, 'bold'),
            pad = ((625, 5), 15)
            )
            ]
    ]
    window = sg.Window(
        '',
        layout,
        text_justification = 'left',
        font = ('Open Sans', 12),
        background_color = COL_2_BACKGROUND_COLOR,
        no_titlebar = True,
        finalize = True
        )
    return window

def cell_map_config():
    '''Handle events for the cell mapping window'''
    window = None
    while True:
        try:
            if window is None:
                prev_page = ''
                prev_menu = ''
                window = cell_map_layout()
                window['-UNASSIGN_MENU-'].update(visible = False)
                button_list = [
                    '-SAVE-',
                    '-MENU_SWAP-',
                    '-PAGE_SWAP-'
                ]
                button_list2 = [
                    '-EXIT-',
                    '-UNASSIGN_MENU-'
                ]
                button_list3 = [
                    '1',
                    '2',
                    '3',
                    '4',
                    '5',
                    '6',
                    'A',
                    'B',
                    'C',
                    'D',
                    'E',
                    'F',
                    'G',
                    'H',
                    'I'
                ]
                for _ in button_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                for _ in button_list2:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                for _ in button_list3:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                window['-TABLE-'].bind('<Enter>', 'ENTER')
                window['-TABLE-'].bind('<Leave>', 'EXIT')
                window['-MMJ_PRODUCT-'].bind('<Enter>', 'ENTER')
                window['-MMJ_PRODUCT-'].bind('<Leave>', 'EXIT')
                locations = menu_locations()
                unassigned_menus, assigned_menus = assigned_menu_locations()
                categories = []
                alias = ''
                for _ in locations.items():
                    if _[0] != '':
                        page = int(_[1][0])
                        menu = _[1][1]
                        mapping = load_mapping(page, menu)
                        name = mapping['MMJ Product']
                        if name != '':
                            categories.append(name)
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '-EXIT-':
                window.close()
                break
            if event == '-TABLE-':
                index = int(values[event][0])
                menu_pos = locations[categories[index]]
                page = int(menu_pos[0])
                for _ in unassigned_menus[page]:
                    window[str(_)].update(button_color = '#FFF on #9B9B9B')
                for _ in assigned_menus[page]:
                    window[str(_)].update(button_color = '#003B48 on #C7F9DC')
                menu = menu_pos[1]
                mapping = load_mapping(page, menu)
                category = mapping['MMJ Product']
                alias = find_alias(category)
                window['-UNIT_PRICE-'].update(str(mapping['Unit Price']))
                window['-BRAND-'].update(str(mapping['Brand']))
                window['-CATEGORY-'].update(str(mapping['Product Category']))
                window['-ROW_START-'].update(int(mapping['First Row Number']))
                window['-ROW_END-'].update(int(mapping['Last Row Number']))
                window['-THC_COL-'].update(str(mapping['THC Column']))
                window['-TYPE_COL-'].update(str(mapping['Type Column']))
                window['-PRODUCT_COL-'].update(str(mapping['Product Column']))
                window['-TITLE_L-'].update(f'Page:\t{page}\nMenu:\t{menu}')
                window['-TITLE_R-'].update(str(alias))
                if prev_page != '' and prev_menu != '':
                    window[str(prev_page)].update(disabled = False,
                                                button_color = '#003B48 on #C7F9DC')
                    window[str(prev_menu)].update(disabled = False,
                                                button_color = '#003B48 on #C7F9DC')
                prev_page = page
                prev_menu = menu
                window[str(page)].update(disabled = True,
                                         button_color = '#003B48 on #59DFAB')
                window[str(menu)].update(disabled = True,
                                         button_color = '#003B48 on #59DFAB')
            if alias == '':
                window['-UNASSIGN_MENU-'].update(visible = False)
            if alias != '':
                window['-UNASSIGN_MENU-'].update(visible = True)
            if event in ['1', '2', '3', '4', '5', '6']:
                for _ in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
                    window[_].update(disabled = False,
                    button_color = '#003B48 on #C7F9DC')
                page = int(event)
                try:
                    for _ in unassigned_menus[page]:
                        window[str(_)].update(button_color = '#FFF on #9B9B9B')
                    for _ in assigned_menus[page]:
                        window[str(_)].update(button_color = '#003B48 on #C7F9DC')
                except KeyError:
                    pass
                mapping = load_mapping(page, menu)
                category = mapping['MMJ Product']
                alias = find_alias(category)
                window['-UNIT_PRICE-'].update(str(mapping['Unit Price']))
                window['-BRAND-'].update(str(mapping['Brand']))
                window['-CATEGORY-'].update(str(mapping['Product Category']))
                window['-ROW_START-'].update(str(mapping['First Row Number']))
                window['-ROW_END-'].update(str(mapping['Last Row Number']))
                window['-THC_COL-'].update(str(mapping['THC Column']))
                window['-TYPE_COL-'].update(str(mapping['Type Column']))
                window['-PRODUCT_COL-'].update(str(mapping['Product Column']))
                window['-TITLE_L-'].update(f'Page: {page}  /  Menu: {menu}')
                window['-TITLE_R-'].update(str(alias))
                if prev_page != '' and prev_menu != '':
                    window[str(prev_page)].update(disabled = False,
                                                  button_color = '#003B48 on #C7F9DC')
                    window[str(prev_menu)].update(disabled = False,
                                                  button_color = '#003B48 on #C7F9DC')
                prev_page = page
                prev_menu = menu
                window[str(page)].update(disabled = True, button_color = '#003B48 on #59DFAB')
                window[str(menu)].update(disabled = True, button_color = '#003B48 on #59DFAB')
                edit_name = False
            if event in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
                menu = event
                mapping = load_mapping(page, menu)
                category = mapping['MMJ Product']
                alias = find_alias(category)
                window['-UNIT_PRICE-'].update(str(mapping['Unit Price']))
                window['-BRAND-'].update(str(mapping['Brand']))
                window['-CATEGORY-'].update(str(mapping['Product Category']))
                window['-ROW_START-'].update(str(mapping['First Row Number']))
                window['-ROW_END-'].update(str(mapping['Last Row Number']))
                window['-THC_COL-'].update(str(mapping['THC Column']))
                window['-TYPE_COL-'].update(str(mapping['Type Column']))
                window['-PRODUCT_COL-'].update(str(mapping['Product Column']))
                window['-TITLE_L-'].update(f'Page: {page}  /  Menu: {menu}')
                window['-TITLE_R-'].update(str(alias))
                if prev_page != '' and prev_menu != '':
                    try:
                        if prev_menu in unassigned_menus[page]:
                            window[str(prev_menu)].update(disabled = False,
                                                        button_color = '#FFF on #9B9B9B')
                        else:
                            window[str(prev_menu)].update(disabled = False,
                                                        button_color = '#003B48 on #C7F9DC')
                    except KeyError:
                        for _ in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
                            window[_].update(disabled = False,
                                                            button_color = '#003B48 on #C7F9DC')
                    window[str(prev_page)].update(disabled = False,
                                                  button_color = '#003B48 on #C7F9DC')
                prev_page = page
                prev_menu = menu
                window[str(page)].update(disabled = True,
                                         button_color = '#003B48 on #59DFAB')
                window[str(menu)].update(disabled = True,
                                         button_color = '#003B48 on #59DFAB')
            if isinstance(event, object):
                for _ in button_list:
                    if event == f'{_}ENTER':
                        window_key = str(event).replace('ENTER', '')
                        disabled = False if window[window_key].Widget['state'] == 'normal' else True
                        if disabled is False:
                            window[f'{_}'].update(
                                button_color = BUTTON_COLOR_HOVER)
                            window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            button_color = BUTTON_COLOR)
                        window.set_cursor('arrow')
                for _ in button_list2:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR)
                        window.set_cursor('arrow')
                for _ in button_list3:
                    if event == f'{_}ENTER':
                        window_key = str(event).replace('ENTER', '')
                        disabled = False if window[window_key].Widget['state'] == 'normal' else True
                        if disabled is False:
                            window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window.set_cursor('arrow')
                if event == '-TABLE-ENTER' or event == '-MMJ_PRODUCT-ENTER':
                    window.set_cursor('hand2')
                if event == '-TABLE-EXIT' or event == '-MMJ_PRODUCT-EXIT':
                    window.set_cursor('arrow')
            if event == '-SAVE-':
                window['-SAVE-'].update(disabled=True)
                window.set_cursor('wait')
                mmj_product = mapping['MMJ Product']
                if values['-MMJ_PRODUCT-']:
                    if values['-MMJ_PRODUCT-'][0] != '':
                        mmj_product = values['-MMJ_PRODUCT-'][0]
                        mmj_product = find_category_name(mmj_product)
                mapping_values = [mmj_product,
                                  values['-UNIT_PRICE-'],
                                  values['-BRAND-'],
                                  values['-CATEGORY-'],
                                  values['-ROW_START-'],
                                  values['-ROW_END-'],
                                  values['-THC_COL-'],
                                  values['-TYPE_COL-'],
                                  values['-PRODUCT_COL-']
                                  ]
                save_mapping(page, menu, mapping, mapping_values)
                window['-SAVE-'].update(disabled=False)
                window.set_cursor('arrow')
                window.close()
                window = None
            if event == '-UNASSIGN_MENU-':
                window['-UNASSIGN_MENU-'].update(disabled=True)
                window.set_cursor('wait')
                unassign_menu(page, menu)
                window['-UNASSIGN_MENU-'].update(disabled=False)
                window.set_cursor('arrow')
                window.close()
                window = None
            if event == '-PAGE_SWAP-':
                move_menu('page')
                window.close()
                window = None
            if event == '-MENU_SWAP-':
                move_menu('menu')
                window.close()
                window = None
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 13)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 13)
                )
    window.close()

def discounts_window():
    '''Creates the layout for discount configuration'''
    sg.theme(THEME)
    alias_list = []
    for _ in [1, 2, 3, 4, 5, 6]:
        page = _
        for _ in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            menu = _
            mapping = load_mapping(page, menu)
            category = mapping['MMJ Product']
            if category != '':
                alias = find_alias(category)
                brand = str(category).split()
                alias_list.append(f'{brand[1]} {alias}')
    column = [
        [text_label(
            _, 33, style=2),
            sg.Input(
                key = f'-{i}-',
                size = (4, 1),
                enable_events = True,
                background_color = COL_1_BACKGROUND_COLOR,
                justification = 'c',
                font = ('Open Sans', 12),
                pad = ((15, 15), 2))
            ] for i, _ in enumerate(alias_list)
         ]
    layout = [
        [sg.Text('Product Category\t\t\t  %',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 pad = (10, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 justification = 'r'
                 )],
        [sg.HSeparator(color = '#F7F9FC')],
        [
        sg.Column(column, scrollable=True,
                  vertical_scroll_only = True,
                  background_color = '#FFF',
                  size = (380, 750))
         ],
        [sg.HSeparator(color = '#F7F9FC')],
        [sg.Text('Overall Discount:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 size = (27, 1),
                 pad = (5, 5),
                 border_width = 0,
                 background_color = '#FFF',
                 justification = 'r'
                 ),
                sg.Input(
                key = '-OVERALL_DISCOUNT-',
                size = (4, 1),
                background_color = COL_1_BACKGROUND_COLOR,
                pad = ((15, 15), 5),
                justification = 'c',
                enable_events = True)
         ],
        [
            sg.Button(
                'Clear All',
                size = (8, 1),
                enable_events = True,
                key = '-CLEAR-',
                button_color = '#62074A on #E15878',
                font = ('Open Sans', 11, 'bold'),
                pad = ((212, 0), (25, 0))
                ),
            sg.Button(
                'Save',
                size = (8, 1),
                enable_events = True,
                key = '-SAVE-',
                button_color = '#003B48 on #30C095',
                font = ('Open Sans', 11, 'bold'),
                pad = ((30, 5), (25, 0))
                ),
            ],
        [sg.Button(
                'Exit',
                size = (8, 1),
                enable_events = True,
                key = '-EXIT-',
                button_color = '#62074A on #E15878',
                font = ('Open Sans', 11, 'bold'),
                pad = ((322, 5), (15, 10))
                )
            ]
        ]
    window = sg.Window(
        'Dispensary Menu Creator: Discounted Products',
        layout,
        icon = PROGRAM_ICON,
        text_justification = 'left',
        font = ('Open Sans', 12),
        background_color = COL_2_BACKGROUND_COLOR,
        no_titlebar = True,
        finalize = True
        )
    return window

def discount_config():
    '''Handle events for the cell mapping window'''
    window = None
    while True:
        try:
            if window is None:
                window = discounts_window()
                categories = category_list()
                discounts = load_discounts()
                button_list = [
                    '-SAVE-'
                ]
                button_list2 = [
                    '-EXIT-',
                    '-CLEAR-'
                ]
                for _ in button_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                for _ in button_list2:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                for i, _ in enumerate(categories):
                    window[f'-{i}-'].update(value = discounts[i])
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '-EXIT-':
                window.close()
                break
            if isinstance(event, object):
                for _ in button_list:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            button_color = BUTTON_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            button_color = BUTTON_COLOR)
                        window.set_cursor('arrow')
                for _ in button_list2:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR)
                        window.set_cursor('arrow')
            if event == '-SAVE-':
                window['-SAVE-'].update(disabled=True)
                window.set_cursor('wait')
                discounts = []
                for i, _ in enumerate(categories):
                    discounts.append(values[f'-{i}-'])
                overall_discount = values['-OVERALL_DISCOUNT-']
                save_discounts(discounts, overall_discount)
                categories = load_categories()
                discounts = load_discounts()
                window['-SAVE-'].update(disabled=False)
                window.set_cursor('arrow')
                window.close()
                window = None
            if event == '-CLEAR-':
                for i, _ in enumerate(categories):
                    window[f'-{i}-'].update(value = '')
            if event == '-EXIT-':
                break
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 13)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 13)
                )
    window.close()

def about_window():
    '''Creates the layout for about()'''
    sg.theme(THEME)
    column_1 = [
        [
        sg.Text(
            'More Info',
            font = ('Open Sans', 13, 'bold'),
            text_color = '#0D70E8',
            pad = (3, 0),
            background_color = COL_1_BACKGROUND_COLOR
            )
            ],
        [
        sg.Image(
            rf'{ICON_GITHUB}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'GitHub Page',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-GITHUB_PAGE-'
            )
            ],
        [
        sg.Image(
            rf'{ICON_CONTACT}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Contact Developer',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-CONTACT_DEV-'
            )
            ],
        [
        sg.Text(
            'UI Imagery',
            font = ('Open Sans', 13, 'bold'),
            text_color = '#0D70E8',
            pad = (3, (25, 0)),
            background_color = COL_1_BACKGROUND_COLOR
            )
            ],
        [
        sg.Image(
            rf'{ICON_ICONS8}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'icons8.com',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-ICONS8-'
            )
            ]
    ]
    column_2 = [
        [
        sg.Text(
            ABOUT,
            background_color = COL_2_BACKGROUND_COLOR
            )
            ],
        [sg.Button(
                'Close',
                size = (6, 1),
                enable_events = True,
                key = '-EXIT-',
                button_color = '#62074A on #E15878',
                font = ('Open Sans', 11, 'bold'),
                pad = ((400, 5), (25, 5))
                )
            ]
        ]
    layout = [
        [
            sg.Column(
                column_1,
                background_color = COL_1_BACKGROUND_COLOR,
                pad = (0, 0),
                size = (185, 300)),
            sg.Column(
                column_2,
                background_color = COL_2_BACKGROUND_COLOR,
                pad = (0, 0),
                size = (475, 300))
            ]
        ]
    window = sg.Window(
        'Dispensary Menu Creator: About',
        layout,
        icon = PROGRAM_ICON,
        text_justification = 'left',
        font = ('Open Sans', 12),
        background_color = COL_1_BACKGROUND_COLOR,
        no_titlebar = True,
        finalize = True
        )
    return window

def about():
    '''Handle events for the about window'''
    window = None
    while True:
        try:
            if window is None:
                window = about_window()
                menu_list = [
                    '-CONTACT_DEV-',
                    '-ICONS8-',
                    '-GITHUB_PAGE-'
                    ]
                button_list = [
                    '-EXIT-'
                    ]
                for _ in menu_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                    window[f'{_}'].bind('<Button-1>', 'CLICK')
                for _ in button_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '-EXIT-':
                break
            if event == '-GITHUB_PAGE-CLICK':
                webbrowser.open(GITHUB_LINK)
            if event == '-CONTACT_DEV-CLICK':
                webbrowser.open('mailto:aaron.hitzeman@gmail.com', new=1)
            if event == '-ICONS8-CLICK':
                webbrowser.open('https://icons8.com/')
            if isinstance(event, object):
                for _ in menu_list:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            text_color = TEXT_LINK_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            text_color = TEXT_LINK_COLOR)
                        window.set_cursor('arrow')
                for _ in button_list:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR)
                        window.set_cursor('arrow')
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 13)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 13)
                )
    window.close()

def categories_window():
    '''Creates the layout for discount configuration'''
    sg.theme(THEME)
    categories_dict = load_categories()
    categories_list = list(categories_dict)
    categories_list.sort()
    alias_list = []
    for _ in categories_list:
        alias = find_alias(_)
        alias_list.append(alias)
    column = [
        [sg.Text(
            'delete',
            key = f'-{i}-',
            size = (6, 1),
            enable_events = True,
            font = ('Open Sans', 11, 'bold'),
            background_color = '#FFF',
            text_color = '#F07B8B',
            pad = (0, 0)),
        text_label(
            _, 33, style=2),
        sg.Text(
            'rename',
            key = f'-{i}_RENAME-',
            size = (7, 1),
            enable_events = True,
            font = ('Open Sans', 11, 'bold'),
            background_color = '#FFF',
            text_color = '#F07B8B',
            pad = ((25, 2), 0)),
        sg.Input(
                alias_list[i],
                key = f'-{i}_ALIAS-',
                size = (33, 1),
                enable_events = True,
                background_color = COL_1_BACKGROUND_COLOR,
                disabled_readonly_background_color = COL_2_BACKGROUND_COLOR,
                justification = 'l',
                border_width = 0,
                disabled = True,
                pad = (5, 0))
            ] for i, _ in enumerate(categories_list)
         ]
    layout = [
        [
        sg.Text('',
                background_color = '#FFF',
                size = (6, 1),
                ),
        sg.Text('Product Category',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 size = (36, 1),
                 pad = (0, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 ),
        sg.Text('Category Alias',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#0D70E8',
                 size = (20, 1),
                 pad = (0, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 ),
            ],
        [sg.HSeparator(color = '#F7F9FC')],
        [
        sg.Column(column, scrollable=True,
                  vertical_scroll_only = True,
                  background_color = '#FFF',
                  size = (770, 755))
            ],
        [sg.HSeparator(color = '#F7F9FC')],
        [sg.Button(
                'Exit',
                size = (8, 1),
                enable_events = True,
                key = '-EXIT-',
                button_color = EXIT_COLOR,
                font = ('Open Sans', 11, 'bold'),
                pad = ((710, 5), (15, 10))
                )
            ]
        ]
    window = sg.Window(
        '',
        layout,
        icon = PROGRAM_ICON,
        text_justification = 'left',
        font = ('Open Sans', 12),
        background_color = COL_2_BACKGROUND_COLOR,
        no_titlebar = True,
        finalize = True
        )
    return window

def categories():
    '''Handle events for the cell mapping window'''
    window = None
    while True:
        try:
            if window is None:
                categories_dict = load_categories()
                categories_list = list(categories_dict)
                categories_list.sort()
                window = categories_window()
                text_list = []
                text_list_2 = []
                disabled = True
                for i, _ in enumerate(categories_list):
                    text_list.append(f'-{i}-')
                    text_list_2.append(f'-{i}_RENAME-')
                button_list = [
                    '-EXIT-'
                ]
                for _ in text_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                    window[f'{_}'].bind('<Button-1>', 'CLICK')
                for _ in text_list_2:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                    window[f'{_}'].bind('<Button-1>', 'CLICK')
                for _ in button_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '-EXIT-':
                window.close()
                break
            if isinstance(event, object):
                for _ in button_list:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            button_color = EXIT_COLOR)
                        window.set_cursor('arrow')
                for _ in text_list:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            text_color = '#FF0000')
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            text_color = '#F07B8B')
                        window.set_cursor('arrow')
                    if event == f'{_}CLICK':
                        category = categories_list[int(event.split('-')[1])]
                        categories_dict.pop(category)
                        save_categories(categories_dict)
                        window.close()
                        window = None
                for _ in text_list_2:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            text_color = '#FF0000')
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            text_color = '#F07B8B')
                        window.set_cursor('arrow')
                    if event == f'{_}CLICK':
                        stop = event.find('_')
                        i = event[1:stop]
                        value = f'-{i}_ALIAS-'
                        if disabled is True:
                            window[value].update(
                                disabled = False)
                            disabled = False
                        elif disabled is False:
                            window[value].update(
                                disabled = True)
                            disabled = True
                            save_alias(categories_list[int(i)], values[value])

        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 13)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 13)
                )
    window.close()

def main_window(update):
    '''Creates the layout for [main_window]'''
    sg.theme(THEME)
    if update is True:
        height = WINDOW_HEIGHT_UPDATE
    else:
        height = WINDOW_HEIGHT
    column_1 = [
        [sg.Text(
            '  Newer Version Available!',
            font = ('Open Sans', 10, 'bold'),
            text_color = '#000',
            size = (22, 1),
            pad = (5, 0),
            border_width = 0,
            background_color = '#75CBF9',
            key = '-UPDATE_MESSAGE-',
            visible = update
            )
        ],
        [
        sg.Text(
            'Menu Creation',
            font = ('Open Sans', 13, 'bold'),
            text_color = '#0D70E8',
            pad = (3, 0),
            background_color = COL_1_BACKGROUND_COLOR
            )
            ],
        [
        sg.Image(
            rf'{ICON_FOLDER}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Saved Menus',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-SAVED_MENUS-'
            )
            ],
        [
        sg.Image(
            rf'{ICON_DISCOUNTED_PRODUCTS}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Discounted Products',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-DISCOUNTED_PRODUCTS-'
            )
            ],
        [
        sg.Text(
            'Menu Configuration',
            font = ('Open Sans', 13, 'bold'),
            text_color = '#0D70E8',
            pad = (3, (25, 0)),
            background_color = COL_1_BACKGROUND_COLOR
            )
            ],
        [
        sg.Image(
            rf'{ICON_PRODUCT_CATEGORIES}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Product Categories',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-PRODUCT_CATEGORIES-'
            )
            ],
        [
        sg.Image(
            rf'{ICON_MENU_ASSIGNMENTS}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Menu Assignments',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-MENU_ASSIGNMENTS-'
            )
            ],
        [
        sg.Text(
            'Look and Feel',
            font = ('Open Sans', 13, 'bold'),
            text_color = '#0D70E8',
            pad = (3, (25, 0)),
            background_color = COL_1_BACKGROUND_COLOR
            )
            ],
        [
        sg.Image(
            rf'{ICON_FOLDER}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Menu Template',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-MENU_TEMPLATE-'
            )
            ],
        [
        sg.Image(
            rf'{ICON_FOLDER}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Menu Logo',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-MENU_LOGO-'
            )
            ],
        [
        sg.Text(
            'Information',
            font = ('Open Sans', 13, 'bold'),
            text_color = '#0D70E8',
            pad = (3, (25, 0)),
            background_color = COL_1_BACKGROUND_COLOR
            )
            ],
        [
        sg.Image(
            rf'{ICON_HELP}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), 0)),
        sg.Text(
            'Help',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), 0),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-HELP-'
            )
            ],
        [
        sg.Image(
            rf'{ICON_ABOUT}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), (0, 25))),
        sg.Text(
            'About',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), (0, 25)),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-ABOUT-'
            )
            ],
        [
        sg.Image(
            rf'{ICON_DOWNLOAD}',
            background_color = COL_1_BACKGROUND_COLOR,
            pad = ((5, 3), (0, 5)),
            key = '-DOWNLOAD_UPDATE_ICON-',
            visible = update
            ),
        sg.Text(
            'Download Update',
            font = ('Open Sans', 12, 'bold'),
            text_color = TEXT_LINK_COLOR,
            pad = ((0, 5), (0, 5)),
            background_color = COL_1_BACKGROUND_COLOR,
            key = '-DOWNLOAD_UPDATE-',
            visible = update
            )
            ],
        [
        sg.Text(
            f'Version:',
            font = ('Open Sans', 11, 'bold'),
            text_color = '#0D70E8',
            pad = (3, 0),
            background_color = COL_1_BACKGROUND_COLOR
            ),
        sg.Text(f'{VERSION}',
                font = ('Open Sans', 11),
                text_color = '#0D70E8',
                pad = (5, 0),
                background_color = COL_1_BACKGROUND_COLOR
            ),
        ]
    ]
    column_2 = [
        [
        sg.Image(
            rf'{PROGRAM_LOGO}',
            background_color = COL_2_BACKGROUND_COLOR,
            pad = (5, (5, 0))
            )
        ],
        [sg.Text(
            '',
            background_color = COL_2_BACKGROUND_COLOR,
            pad = (0, 0)
            )
        ],
        [
        sg.Text(
            'Please select an exported inventory file → ',
            key = '-FILE_NAME-',
            enable_events = True,
            background_color = '#FFF',
            pad = ((5, 0), 0),
            justification = 'r',
            size = (40, 1),
            font = ('Open Sans', 12, 'bold')
            ),
         sg.FileBrowse(
            key='-FILE_BROWSE-',
            enable_events = True,
            pad = (11, 0),
            button_color = BUTTON_COLOR,
            font = ('Open Sans', 11, 'bold'),
            file_types = (('Current Inventory', '.xlsx'),)
            )
        ],
        [
        sg.ProgressBar(
            100,
            orientation = 'h',
            size = (36.7, 10),
            key = '-P_BAR-',
            border_width = 0,
            pad = ((5, 5), 0),
            bar_color = '#0D70E8 on #333',
            visible = False
            )
        ],
        [sg.Text('0.0 / 100.0%',
                 background_color = '#FFF',
                 enable_events = True,
                 key = '-P_PERCENT-',
                 pad = (0, 0),
                 size = (51, 1),
                 font = ('Open Sans', 10, 'bold'),
                 justification = 'r',
                 visible = False
                 )],
        [sg.Button('Create Menu',
                   size = (11, 1),
                   enable_events = True,
                   button_color = '#003B48 on #30C095',
                   font = ('Open Sans', 11, 'bold'),
                   key = '-CREATE_MENU-',
                   pad = ((379, 5), 20),
                   visible = False
                   )],
        [
        sg.Text('',
                background_color = COL_2_BACKGROUND_COLOR,
                size = (36, 1),
                key = '-UNASSIGNED_CATEGORIES-',
                justification = 'r',
                font = ('Open Sans', 11, 'bold'),
                text_color = '#0D70E8',
                pad = (0, (WINDOW_HEIGHT - 180, 0))
                ),
        sg.Text(
            'unassigned categories',
            background_color = COL_2_BACKGROUND_COLOR,
            justification = 'r',
            font = ('Open Sans', 11),
            text_color = '#0D70E8',
            pad = (0, (WINDOW_HEIGHT - 180, 0))
            )
            ]
    ]
    layout = [
        [
            sg.Column(
                column_1,
                background_color = COL_1_BACKGROUND_COLOR,
                pad = (0, (0, 5)),
                size = (200, height)),
            sg.Column(
                column_2,
                background_color = COL_2_BACKGROUND_COLOR,
                pad = (0, (0, 5)),
                size = (500, height))
            ]
        ]
    window = sg.Window(
        'Dispensary Menu Creator',
        layout,
        icon = PROGRAM_ICON,
        text_justification = 'left',
        font = ('Open Sans', 13),
        background_color = COL_1_BACKGROUND_COLOR,
        finalize = True
        )
    return window

def main():
    '''Handle events for the main window'''
    window = None
    while True:
        try:
            if window is None:
                if AVAILABLE_UPDATE is True:
                    update = True
                else:
                    update = False
                window = main_window(update)
                unassigned_cat_list = unassigned_categories()
                num_unassigned_categories = len(unassigned_cat_list)
                window['-UNASSIGNED_CATEGORIES-'].update(num_unassigned_categories)
                window.Refresh()
                menu_list = [
                    '-SAVED_MENUS-',
                    '-DISCOUNTED_PRODUCTS-',
                    '-PRODUCT_CATEGORIES-',
                    '-MENU_ASSIGNMENTS-',
                    '-MENU_TEMPLATE-',
                    '-MENU_LOGO-',
                    '-HELP-',
                    '-ABOUT-',
                    '-DOWNLOAD_UPDATE-'
                ]
                button_list = [
                    '-FILE_BROWSE-',
                    '-CREATE_MENU-'
                ]
                for _ in menu_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
                    window[f'{_}'].bind('<Button-1>', 'CLICK')
                for _ in button_list:
                    window[f'{_}'].bind('<Enter>', 'ENTER')
                    window[f'{_}'].bind('<Leave>', 'EXIT')
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Exit':
                break
            if isinstance(event, object):
                for _ in menu_list:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            text_color=TEXT_LINK_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            text_color=TEXT_LINK_COLOR)
                        window.set_cursor('arrow')
                    if event == f'{_}CLICK':
                        if event == '-SAVED_MENUS-CLICK':
                            path = pathlib.PurePath(MAIN_DIRECTORY,
                                                    'saved_menus')
                            os.startfile(path)
                        if event == '-DISCOUNTED_PRODUCTS-CLICK':
                            window.minimize()
                            discount_config()
                            window.normal()
                        if event == '-PRODUCT_CATEGORIES-CLICK':
                            window.minimize()
                            categories()
                            window.normal()
                        if event == '-MENU_ASSIGNMENTS-CLICK':
                            window.minimize()
                            cell_map_config()
                            window.normal()
                        if event == '-MENU_TEMPLATE-CLICK':
                            path = pathlib.PurePath(MAIN_DIRECTORY,
                                                    'config_files',
                                                    'menu_template')
                            os.startfile(path)
                        if event == '-MENU_LOGO-CLICK':
                            path = pathlib.PurePath(MAIN_DIRECTORY,
                                                    'img',
                                                    'menu')
                            os.startfile(path)
                        if event == '-HELP-CLICK':
                            path = pathlib.PurePath(MAIN_DIRECTORY,
                                                    'help_files',
                                                    'Dispensary Menu Creator.pdf')
                            os.startfile(path)
                        if event == '-ABOUT-CLICK':
                            window.minimize()
                            about()
                            window.normal()
                        if event == '-DOWNLOAD_UPDATE-CLICK':
                            webbrowser.open(DOWNLOAD_LINK)
                if event == '-CREATE_MENU-':
                    window['-FILE_NAME-'].update(background_color = '#FFF')
                    window['-FILE_NAME-'].update('Creating the menu...')
                    window['-CREATE_MENU-'].update(disabled=True)
                    window.set_cursor('wait')
                    exported_packages = pathlib.PurePath(values['-FILE_BROWSE-'])
                    current_packages = pd.read_excel(exported_packages,
                                                    sheet_name = 'All Packages')
                    full_menu = build_menu(current_packages)
                    np.sort(full_menu['category'].unique())
                    save_all(full_menu, window)
                    window['-CREATE_MENU-'].update(disabled=False)
                    window.set_cursor('arrow')
                    values['-FILE_BROWSE-'] = ''
                for _ in button_list:
                    if event == f'{_}ENTER':
                        window[f'{_}'].update(
                            button_color= BUTTON_COLOR_HOVER)
                        window.set_cursor('hand2')
                    if event == f'{_}EXIT':
                        window[f'{_}'].update(
                            button_color = BUTTON_COLOR)
                        window.set_cursor('arrow')
            if values['-FILE_BROWSE-'] != '':
                window['-P_BAR-'].update(
                    visible = True
                    )
                window['-FILE_NAME-'].update(
                    background_color = COL_1_BACKGROUND_COLOR
                    )
                window['-P_PERCENT-'].update(
                    visible = True
                    )
                window['-CREATE_MENU-'].update(
                    visible = True
                    )
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 13)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 13)
                )
    window.close()

if __name__ == '__main__':
    main()
