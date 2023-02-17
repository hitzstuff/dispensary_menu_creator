import shutil
from datetime import date
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

MAIN_DIRECTORY = str(pathlib.Path( __file__ ).parent.absolute())
PROGRAM_ICON = str(MAIN_DIRECTORY) + r'\img\menu.ico'
MENU_LOGO = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'menu_logo.png')
PROGRAM_LOGO = pathlib.PurePath(MAIN_DIRECTORY, 'img', 'program_logo.png')
THEME = 'SystemDefault1'
PROGRAM_NAME = 'Dispensary Menu Creator'
GITHUB_LINK = 'https://github.com/hitzstuff/dispensary_menu_creator'
CATEGORIES_FILE = pathlib.PurePath(MAIN_DIRECTORY, 'config', 'categories.cfg')

MENU_BAR = [
    ['File',
        ['Exit']
        ],
    ['Edit',
        ['Menu Mapping Configuration', 'Discounted Products']
        ],
    ['Help',
        ['GitHub Page', 'Check for Updates', 'About']
        ]
    ]

# Version checking
request = requests.get(GITHUB_LINK, timeout=5)
parse = bs4.BeautifulSoup(request.text, 'html.parser')
parse_part = parse.select('div#readme p')
NEWEST_VERSION = str(parse_part[0])
NEWEST_VERSION = NEWEST_VERSION.split()[-1][:-4]
DOWNLOAD_LINK = str(parse_part[1])
DOWNLOAD_LINK = ((DOWNLOAD_LINK.split()[-1][:-4]).split('>')[1]).split('<')[0]

# Preformatted string for the 'About' menu
ABOUT = (
    f'Current Version:\n{VERSION}\n\n' +
    f'Latest Version:\n{NEWEST_VERSION}\n\n'
    +
    'Developed by:\nAaron Hitzeman\n' +
    'aaron.hitzeman@gmail.com\n\n'
    +
    'Visit the GitHub page for more information.'
)

# Preformatted string for the 'About' menu when an update is available
ABOUT_UPDATE = (
    f'Current Version:\n{VERSION}\n\n' +
    f'Latest Version:\n{NEWEST_VERSION}\n\n'
    +
    'A newer version of this program is available!' +
    '  Please use the "Check for Updates" button to download the latest version.\n\n'
    +
    'Developed by:\nAaron Hitzeman\n' +
    'aaron.hitzeman@gmail.com\n\n'
    +
    'Visit the GitHub page for more information.'
)

# Preformatted update messages
UPDATE_MSG_YES = (
    f'Current Version:\n{VERSION}\n\n' +
    f'Latest Version:\n{NEWEST_VERSION}\n\n'
    +
    'A newer version of this program is available!  Would you like to download it?'
)
UPDATE_MSG_NO = (
    f'Current Version:\n{VERSION}\n\n' +
    f'Latest Version:\n{NEWEST_VERSION}\n\n'
    +
    'There are currently no updates available.'
)

def update_check():
    '''Checks the program's version against the one on the GitHub page'''
    if VERSION != NEWEST_VERSION:
        status = True
    else:
        status = False
    return status

def create_menu_file():
    '''Creates a new menu from the template file'''
    menu_template = pathlib.PurePath(MAIN_DIRECTORY, 'config', 'menu_template.xlsx')
    today = date.today()
    month = today.strftime('%m')
    day = today.strftime('%d')
    year = today.strftime('%Y')
    name = f'Menu {month}-{day}-{year}.xlsx'
    file_path = pathlib.PurePath(MAIN_DIRECTORY, '_schedules', name)
    shutil.copy(menu_template, file_path)
    menu_file = openpyxl.load_workbook(file_path)
    return menu_file, file_path

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
    # DELETE_ME: if nothing breaks next time EPO stuff comes in...
    #dataframe[dataframe['sku_retail_display_name'].str.contains('EPO') == False]
    return dataframe

def new_categories(dataframe):
    '''Combines unit_price and category to create a new category'''
    categories = []
    for i in range(len(dataframe)):
        unit_price = dataframe.unit_price.iloc()[i]
        brand = category = dataframe.brand.iloc()[i]
        category = dataframe.category.iloc()[i]
        cat_df = dataframe[dataframe.category == category]
        new_category_a = f'${unit_price} {category}'
        new_category_b = f'{brand} {category}'
        num_prices = len(cat_df.unit_price.unique())
        num_brands = len(cat_df.brand.unique())
        if num_prices > 1:
            if num_brands == 1:
                if new_category_a not in categories:
                    categories.append(new_category_a)
            else:
                if new_category_b not in categories:
                    categories.append(new_category_b)
        else:
            if category not in categories:
                categories.append(category)
    return categories

def coral_reefer_fix(full_menu):
    '''Due to incorrect inventory labeling, this is required to fix certain categories'''
    # This is actually a dictionary, but shit dictionary didn't sound as cool
    shit_list = {'Island Time': 'Indica',
                "Sunset Sailin'": 'Hybrid',
                "Surfin' in a Hurricane": 'Sativa'
                }
    for i, _ in enumerate(full_menu.strain):
        if _ in shit_list:
            full_menu.product_type.iloc[i] = shit_list[_]
    return full_menu

def build_menu(dataframe):
    '''Organizes a DataFrame into a structure suitable for a menu'''
    cleaned_packages = df_clean(dataframe)
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
    sku = dataframe['sku_retail_display_name']
    brand = dataframe['brand']
    size = dataframe['product_size']
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
    for i, _ in enumerate(dataframe['product_size']):
        if 'mg' in _:
            dataframe['thc'][i] = str(_)
            if 'ct' in sku[i].split()[-1]:
                dataframe['product_size'][i] = sku[i].split()[-1]
            else:
                dataframe['product_size'][i] = '1ct'
    price = dataframe['unit_price']
    for i, _ in enumerate(dataframe['category']):
        dataframe['category'][i] = f'{price[i]} {brand[i]} {size[i]} {_}'
    for i, _ in enumerate(dataframe['strain']):
        if str(_) == 'nan' or str(_) == '':
            dataframe['strain'][i] = str(dataframe['product_type'][i]) + ' Blend'
        if ('HYB' in str(_)[0:4]) or ('SAT' in str(_)[0:4]) or ('IND' in str(_)[0:4]):
            dataframe['strain'][i] = str(_)[4:]
        if 'THC' == _:
            sku = dataframe['sku_retail_display_name'][i]
            new_sku = ' '.join(sku.split()[2:])
            stop = new_sku.find('(')
            new_strain = ' '.join(new_sku[:stop].split()[:-1])
            dataframe['strain'][i] = new_strain
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
    dataframe = coral_reefer_fix(dataframe)
    dataframe = dataframe.query('brand != "EPO"').reset_index(drop=True)
    dataframe.sort_values(by = ['category', 'strain'], inplace=True)
    dataframe.sort_values(by = ['unit_price', 'thc', 'product_type'], ascending=False, inplace=True)
    return dataframe

def save_product_type(sheet, menu_category, column, first_row, last_row):
    '''Saves the product types to the menu'''
    prod_type = []
    for _ in list(menu_category['product_type']):
        prod_type.append(_)
    row = first_row
    for _ in prod_type:
        cell = f'{column}{row}'
        sheet[cell].value = _
        if row <= last_row:
            row += 1
            cell = f'{column}{row}'
    while row <= last_row:
        cell = f'{column}{row}'
        sheet[cell].value = ''
        row += 1
    return None

def save_product_thc(sheet, menu_category, column, first_row, last_row):
    '''Saves the THC % to the menu'''
    thc = []
    for _ in list(menu_category['thc']):
        thc.append(_)
    row = first_row
    for _ in thc:
        cell = f'{column}{row}'
        sheet[cell].value = _
        if row <= last_row:
            row += 1
            cell = f'{column}{row}'
    while row <= last_row:
        cell = f'{column}{row}'
        sheet[cell].value = ''
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
        sheet[cell].value = _
        if row <= last_row:
            row += 1
            cell = f'{column}{row}'
    while row <= last_row:
        cell = f'{column}{row}'
        sheet[cell].value = ''
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
            if sale_percent == 0:
                unit_price = float(list(menu['unit_price'])[0])
                discount_msg = ''
                sheet[discount_cell].value = discount_msg
            if sale_percent != 0:
                unit_price = float(list(menu['unit_price'])[0]) * (1 - (sale_percent / 100))
                discount_msg = f'SALE: {sale_percent}% OFF!'
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
        if sale_percent != 0 or sale_percent != '':
            sheet[unit_price_pos].font = openpyxl.styles.Font(
                'Bahnschrift',
                size=18,
                color='FF0000',
                b=True
                )
        if sale_percent == 0 or sale_percent == '':
            sheet[unit_price_pos].font = openpyxl.styles.Font(
                'Bahnschrift',
                size=18,
                color='000000',
                b=True
                )
        sheet[category_pos].value = alias
        sheet[brand_pos].value = product_brand
        # Save the new values to the workbook
        workbook.save(workbook_path)
        workbook.close()
    return None

def save_all(full_menu):
    '''Saves each menu to the Excel workbook'''
    workbook, workbook_path = create_menu_file()
    pages = [1, 2, 3, 4, 5]
    menus = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    for _ in pages:
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
    return None

def create_window(title, layout, program_icon=None, background_color='#FFF'):
    '''Creates a PySimpleGUI window'''
    window = sg.Window(
        title,
        layout,
        icon = program_icon,
        text_justification = 'left',
        font = ('Open Sans', 12),
        background_color = background_color,
        finalize = True
        )
    return window

def mapping_file(page, menu_position):
    '''Returns the mapping file for a specified page and menu position'''
    file = pathlib.PurePath(
            MAIN_DIRECTORY,
            'config',
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

def save_categories(categories):
    '''Saves all category information to the categories.cfg file'''
    file = CATEGORIES_FILE
    with open(file, 'w', encoding='UTF-8') as file:
        json_dump(categories, file)
    return None

def populate_categories(menu):
    '''Populates categories from a new menu and adds them to the current dictionary'''
    category_list = list(menu.category.unique())
    categories = load_categories()
    for _ in category_list:
        if _ not in categories:
            categories[_] = [_, 0, '']
    save_categories(categories)
    return categories

def menu_locations(cat_type=None):
    '''Returns a list of menus and their product categories'''
    menu_dict = {}
    for _ in [1, 2, 3, 4, 5]:
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

def category_list():
    '''Collects category names from the worksheet cell mapping files'''
    cat_list = []
    for _ in [1, 2, 3, 4, 5]:
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
            categories = json_load(file)
    except FileNotFoundError as error:
        sg.popup(f'exception {error}\n\nNo categories file found.', title='', font = ('Open Sans', 12))
    return categories

def save_discounts(discount_values, overall_discount):
    '''Saves the discount values to their respective categories in the categories.cfg file'''
    categories = load_categories()
    cat_list = category_list()
    for i, _ in enumerate(cat_list):
        if _ != '':
            if discount_values[i] == '':
                categories[_][1] = overall_discount
            else:
                categories[_][1] = discount_values[i]
    file = CATEGORIES_FILE
    # Opens the discount file and overwrites it with the new value
    with open(file, 'w', encoding='UTF-8') as file:
        json_dump(categories, file)
    return None

def load_discounts():
    '''Loads the discount values from the categories.cfg file'''
    categories = load_categories()
    cat_list = category_list()
    discounts = []
    for _ in cat_list:
        discounts.append(categories[_][1])
    return discounts

def find_discount(category):
    '''Given a product category, returns the current discount percent'''
    categories = load_categories()
    discount = categories[category][1]
    if discount == '':
        discount = 0
    discount = int(discount)
    return discount

def save_menu_pos(menu_positions):
    '''Saves the menu positions to their respective categories in the categories.cfg file'''
    categories = load_categories()
    for i, _ in enumerate(categories):
        categories[_][1] = menu_positions[i]
    file = CATEGORIES_FILE
    # Opens the discount file and overwrites it with the new value
    with open(file, 'w', encoding='UTF-8') as file:
        json_dump(categories, file)
    return None

def find_menu_pos(category):
    '''Given a product category, returns the menu position'''
    categories = load_categories()
    menu_position = categories[category][3]
    return menu_position

def find_category(page, menu):
    '''Given a page and a menu position, returns the category'''
    position = f'{page}{menu}'
    categories = load_categories()
    for _ in categories:
        if categories[_][2] == position:
            category = _
    return category

def find_alias(category):
    '''Given a product category, returns its alias'''
    categories = load_categories()
    try:
        alias = categories[category][0]
    except KeyError:
        alias = category
    return alias

def find_category_name(alias):
    '''Given a product alias, returns its category name'''
    categories = load_categories()
    category_dict = {}
    for _ in categories:
        key = categories[_][0]
        value = _
        category_dict[key] = value
    return category_dict[alias]

def save_alias(category, alias):
    '''Given a category name and its alias, saves them to categories.cfg'''
    categories = load_categories()
    categories[category][0] = alias
    save_categories(categories)
    return None

def range_char(start, stop):
    '''Equivalent to Python's built-in range function, but for letters instead of numbers'''
    converted_characters = (chr(_) for _ in range(ord(start), ord(stop) + 1))
    return converted_characters

def text_label(text, width):
    '''Returns a text label'''
    label = sg.Text(text+': ',
                    justification='l',
                    size =  (width, 1),
                    background_color = '#FFF',
                    text_color = '#000',
                    pad = ((5, 0), 2),
                    font = ('Open Sans', 12, 'bold')
                    )
    return label

def textButton(text, background_color, text_color, style=1):
    '''Given text, returns an element that looks like a button'''
    if style == 1:
        button = sg.Text(text,
                        key = ('-B-', text),
                        enable_events = True,
                        justification = 'r',
                        background_color = background_color,
                        font = ('Open Sans', 12, 'underline'),
                        pad = (5, 0),
                        text_color = text_color)
    if style == 2:
        button = sg.Text(text,
                        key = ('-B-', text),
                        relief = 'raised',
                        enable_events = True,
                        background_color = background_color,
                        text_color = text_color)
    return button

def bind_button(window, button_text):
    '''Magic code that enables mouseover highlighting to work'''
    _ = button_text
    window[('-B-', _)].bind('<Enter>', 'ENTER')
    window[('-B-', _)].bind('<Leave>', 'EXIT')
    return None

def discounts_layout():
    '''Creates the layout for discount configuration'''
    sg.theme(THEME)
    alias_list = []
    for _ in [1, 2, 3, 4, 5]:
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
            _, 30),
            sg.Input(
                key = f'-{i}-',
                size = (4, 1),
                enable_events = True,
                background_color = '#E3FFFA',
                justification = 'c',
                pad = ((15, 15), 2))
            ] for i, _ in enumerate(alias_list)
         ]
    layout = [
        [sg.Text('Product Category\t\t             %',
                 font = ('Open Sans', 14, 'bold'),
                 text_color = '#009678',
                 pad = (10, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 justification = 'r'
                 )],
        [sg.HSeparator(color = '#E3FFFA')],
        [
        sg.Column(column, scrollable=True, vertical_scroll_only = True, background_color = '#FFF')
         ],
        [sg.HSeparator(color = '#E3FFFA')],
        [sg.Text('Overall Discount:',
                 font = ('Open Sans', 12, 'bold'),
                 text_color = '#009678',
                 size = (30, 1),
                 pad = (5, 5),
                 border_width = 0,
                 background_color = '#FFF',
                 justification = 'r'
                 ),
                sg.Input(
                key = '-OVERALL_DISCOUNT-',
                size = (4, 1),
                background_color = '#E3FFFA',
                pad = ((15, 15), 5),
                justification = 'c',
                enable_events = True)
         ],
        [
            sg.Button(
                'Clear All',
                size = (10, 1),
                enable_events = True,
                key = '-CLEAR-',
                button_color = '#FFF on #FF0000',
                pad = ((170, 5), (30, 15))
                ),
            sg.Button(
                'Save',
                size = (10, 1),
                enable_events = True,
                key = '-SAVE-',
                button_color = '#FFF on #009678',
                pad = (5, (30, 15))
                ),
            ],
        [sg.Button(
                'Exit',
                size = (10, 1),
                enable_events = True,
                key = '-EXIT-',
                button_color = '#FFF on #F08080',
                pad = ((280, 15), (0, 15))
                )
            ]
        ]
    return layout

def discount_config():
    '''Handle events for the cell mapping window'''
    window = None
    while True:
        try:
            if window is None:
                layout = discounts_layout()
                window = create_window('Discount Configuration', layout, PROGRAM_ICON)
                categories = category_list()
                discounts = load_discounts()
                for i, _ in enumerate(categories):
                    window[f'-{i}-'].update(value = discounts[i])
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '-EXIT-':
                window.close()
                break
            if event == '-SAVE-':
                discounts = []
                for _ in range(0, len(categories)):
                    discounts.append(values[f'-{_}-'])
                overall_discount = values['-OVERALL_DISCOUNT-']
                save_discounts(discounts, overall_discount)
                categories = load_categories()
                discounts = load_discounts()
                for i, _ in enumerate(categories):
                    window[f'-{i}-'].update(value = discounts[i])
                sg.popup('Discounts were successfully saved.', title='', font = ('Open Sans', 12))
            if event == '-CLEAR-':
                for i in range(0, len(categories)):
                    window[f'-{i}-'].update(value = '')
            if event == '-EXIT-':
                break
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 12)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 12)
                )
    window.close()

def table_categories():
    '''Populates a list of categories to act as values for a PySimpleGUI table'''
    category_names = load_categories()
    category_list = []
    for _ in category_names:
        category_list.append(category_names[_][0])
    locations = menu_locations()
    categories = []
    for _ in locations:
        page = int(locations[_][0])
        menu = locations[_][1]
        mapping = load_mapping(page, menu)
        name = mapping['MMJ Product']
        alias = find_alias(name)
        brand = str(name).split()
        if name != '':
            categories.append(f'"{page} {menu}\t{brand[1]}\t{alias}"')
    categories = np.array(categories)
    np.reshape(categories, (len(categories), 1))
    return categories, category_list

def unassign_menu(page, menu):
    '''Given a page number and menu letter, replaces the "MMJ Product" value with a blank value'''
    mapping = load_mapping(page, menu)
    mapping['MMJ Product'] = ''
    save_mapping(page, menu, mapping, None)
    return None

def cell_map_layout():
    '''Layout for worksheet cell mapping'''
    categories, category_list = table_categories()
    sg.theme(THEME)
    menu_category_column = [
        [sg.Table(values = categories,
                  headings = ['Menu\tBrand\t\tCategory\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t'],
                  justification = 'left',
                  key = '-TABLE-',
                  enable_events = True,
                  size = (len(categories), len(categories)),
                  background_color = '#FFF',
                  text_color = '#000',
                  font = ('Open Sans', 12),
                  selected_row_colors = '#FFF on #009678',
                  header_background_color = '#009678',
                  header_text_color = 'yellow',
                  pad = 0,
                  hide_vertical_scroll = True,
                  max_col_width = 25,
                  select_mode = 'extended',
                  auto_size_columns = True
                  )]]
    menu_mapping_column = [
        [sg.Text('New Category:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#009678',
                 pad = (5, 0),
                 border_width = 0,
                 background_color = '#FFF'
                 )],
        [sg.Listbox(values = category_list,
                 key = '-MMJ_PRODUCT-',
                 enable_events = True,
                 font = ('Open Sans', 12),
                 text_color = '#000',
                 background_color = '#E3FFFA',
                 highlight_background_color = '#009678',
                 highlight_text_color = '#FFF',
                 pad = (5, 0),
                 size = (33, 6)
                 )],
        [sg.Text('', background_color = '#FFF')],
        [sg.Text('Page Number:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#009678',
                 pad = (5, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 )],
        [sg.Button(f'{num}',
                   key = f'{num}',
                   enable_events = True,
                   size = (3, 1),
                   button_color = '#FFF on #333',
                   disabled_button_color = 'yellow on #009678'
                   ) for num in range(1, 6)
         ],
        [sg.Text('', background_color = '#FFF')],
        [sg.Text('Menu Position:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#009678',
                 pad = (5, 0),
                 border_width = 0,
                 background_color = '#FFF',
                 )],
        [sg.Button(f'{letter}',
                   key = f'{letter}',
                   enable_events = True,
                   size = (3, 1),
                   button_color = '#FFF on #333',
                   disabled_button_color = 'yellow on #009678'
                   ) for letter in range_char('A', 'C')
         ],
        [sg.Button(f'{letter}',
                   key = f'{letter}',
                   enable_events = True,
                   size = (3, 1),
                   button_color = '#FFF on #333',
                   disabled_button_color = 'yellow on #009678'
                   ) for letter in range_char('D', 'F')
         ],
        [sg.Button(f'{letter}',
                   key = f'{letter}',
                   enable_events = True,
                   size = (3, 1),
                   button_color = '#FFF on #333',
                   disabled_button_color = 'yellow on #009678'
                   ) for letter in range_char('G', 'I')
         ],
        [sg.Text('', background_color = '#FFF')],
        [sg.Text('Worksheet Cell Mapping:',
                 font = ('Open Sans', 13, 'bold'),
                 text_color = '#009678',
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
                background_color = '#E3FFFA',
                pad = ((0, 15), 2))
            ],
        [text_label(
        'Brand', 16),
        sg.Input(
            key = '-BRAND-',
            justification = 'c',
            size = (4, 1),
            enable_events = True,
            background_color = '#E3FFFA',
            pad = ((0, 15), 2))
        ],
        [text_label(
            'Product Category', 16),
            sg.Input(
                key = '-CATEGORY-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#E3FFFA',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'First Row Number', 16),
            sg.Input(
                key = '-ROW_START-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#E3FFFA',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'Last Row Number', 16),
            sg.Input(
                key = '-ROW_END-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#E3FFFA',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'THC Column', 16),
            sg.Input(
                key = '-THC_COL-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#E3FFFA',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'Type Column', 16),
            sg.Input(
                key = '-TYPE_COL-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#E3FFFA',
                pad = ((0, 15), 2))
            ],
        [text_label(
            'Product Column', 16),
            sg.Input(
                key = '-PRODUCT_COL-',
                justification = 'c',
                size = (4, 1),
                enable_events = True,
                background_color = '#E3FFFA',
                pad = ((0, 15), 2))
            ],
        [sg.Button(
                'Save',
                size = (10, 1),
                enable_events = True,
                key = '-SAVE-',
                button_color = '#FFF on #009678',
                pad = ((220, 5), (40, 15))
                )
            ],
        [sg.Button(
            'Exit',
            size = (10, 1),
            enable_events = True,
            key = '-EXIT-',
            button_color = '#FFF on #F08080',
            pad = ((220, 5), (0, 15))
            )
            ]
        ]
    title_column_l = [
        [sg.Text('',
                 enable_events = True,
                 key = '-TITLE_L-',
                 background_color = '#FFF',
                 text_color = '#009678',
                 justification = 'left',
                 pad = ((0, 5), (5, 0)),
                 font = ('Open Sans', 14, 'bold')
                 )],
        [sg.Button('Unassign Menu',
                   size = (16, 1),
                   enable_events = True,
                   key = '-UNASSIGN_MENU-',
                   button_color = '#FFF on #F08080',
                   font = ('Open Sans', 10),
                   pad = (5, (10, 20))
        )]
    ]
    title_column_r = [
        [sg.Text('',
                 enable_events = True,
                 key = '-TITLE_R-',
                 background_color = '#FFF',
                 text_color = '#009678',
                 justification = 'left',
                 pad = (5, (5, 0)),
                 font = ('Open Sans', 14, 'bold')
                 )],
        [textButton('edit name',
                 background_color = '#FFF',
                 text_color = '#009678',
                 style = 1
                 ),
                 sg.Input(size = (33, 1),
                  enable_events = True,
                  key = '-CATEGORY_ALIAS-',
                  pad = ((5, 0), 0),
                  background_color = '#E3FFFA',
                  disabled_readonly_background_color = '#FFF',
                  disabled = True,
                  visible = False
                  )
                 ]
    ]
    layout = [
        [sg.Column(title_column_l,
                   background_color = '#FFF',
                   size = (230, 90),
                   pad = ((10, 20), (5, 0))
                   ),
         sg.Column(title_column_r,
                   background_color = '#FFF',
                   justification = 'left',
                   size = (450, 90),
                   pad = ((10, 10), (5, 0))
                   )
         ],
        [sg.Column(menu_category_column,
                   pad = ((0, 10), 10),
                   vertical_alignment = 'top'),
         sg.Column(menu_mapping_column,
                   pad = (10, 10),
                   vertical_alignment = 'top',
                   background_color = '#FFF'),
         ]
    ]
    return layout

def cell_map_config():
    '''Handle events for the cell mapping window'''
    window = None
    while True:
        try:
            if window is None:
                prev_page = ''
                prev_menu = ''
                edit_name = False
                layout = cell_map_layout()
                window = create_window('Cell Mapping Configuration', layout, PROGRAM_ICON)
                window[('-B-', 'edit name')].update(visible = False)
                window['-UNASSIGN_MENU-'].update(visible = False)
                bind_button(window, 'edit name')
                locations = menu_locations()
                categories = []
                alias = ''
                for _ in locations:
                    if _ != '':
                        page = int(locations[_][0])
                        menu = locations[_][1]
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
                menu = menu_pos[1]
                mapping = load_mapping(page, menu)
                category = mapping['MMJ Product']
                alias = find_alias(category)
                window['-CATEGORY_ALIAS-'].update(str(alias))
                window['-UNIT_PRICE-'].update(str(mapping['Unit Price']))
                window['-BRAND-'].update(str(mapping['Brand']))
                window['-CATEGORY-'].update(str(mapping['Product Category']))
                window['-ROW_START-'].update(int(mapping['First Row Number']))
                window['-ROW_END-'].update(int(mapping['Last Row Number']))
                window['-THC_COL-'].update(str(mapping['THC Column']))
                window['-TYPE_COL-'].update(str(mapping['Type Column']))
                window['-PRODUCT_COL-'].update(str(mapping['Product Column']))
                window['-TITLE_L-'].update(f'Page: {page}  /  Menu: {menu}')
                window['-TITLE_R-'].update(str(alias))
                if prev_page != '' and prev_menu != '':
                    window[str(prev_page)].update(disabled = False, button_color = '#FFF on #333')
                    window[str(prev_menu)].update(disabled = False, button_color = '#FFF on #333')
                prev_page = page
                prev_menu = menu
                window[str(page)].update(disabled = True, button_color = 'yellow on #009678')
                window[str(menu)].update(disabled = True, button_color = 'yellow on #009678')
                window['-CATEGORY_ALIAS-'].update(disabled = True, visible = False)
                edit_name = False
            if alias == '':
                window[('-B-', 'edit name')].update(visible = False)
                window['-UNASSIGN_MENU-'].update(visible = False)
            if alias != '':
                window[('-B-', 'edit name')].update(visible = True)
                window['-UNASSIGN_MENU-'].update(visible = True)
            if event == '-EDIT_NAME-':
                if edit_name is False:
                    window['-CATEGORY_ALIAS-'].update(disabled = False, visible = True)
                    edit_name = True
                else:
                    window['-CATEGORY_ALIAS-'].update(disabled = True, visible = False)
                    edit_name = False
            if event in ['1', '2', '3', '4', '5']:
                page = int(event)
                mapping = load_mapping(page, menu)
                category = mapping['MMJ Product']
                alias = find_alias(category)
                window['-CATEGORY_ALIAS-'].update(str(alias))
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
                    window[str(prev_page)].update(disabled = False, button_color = '#FFF on #333')
                    window[str(prev_menu)].update(disabled = False, button_color = '#FFF on #333')
                prev_page = page
                prev_menu = menu
                window[str(page)].update(disabled = True, button_color = 'yellow on #009678')
                window[str(menu)].update(disabled = True, button_color = 'yellow on #009678')
                window['-CATEGORY_ALIAS-'].update(disabled = True)
                edit_name = False
            if event in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
                menu = event
                mapping = load_mapping(page, menu)
                category = mapping['MMJ Product']
                alias = find_alias(category)
                window['-CATEGORY_ALIAS-'].update(str(alias))
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
                    window[str(prev_page)].update(disabled = False, button_color = '#FFF on #333')
                    window[str(prev_menu)].update(disabled = False, button_color = '#FFF on #333')
                prev_page = page
                prev_menu = menu
                window[str(page)].update(disabled = True, button_color = 'yellow on #009678')
                window[str(menu)].update(disabled = True, button_color = 'yellow on #009678')
                window['-CATEGORY_ALIAS-'].update(disabled = True)
                edit_name = False
            if isinstance(event, tuple):
                if event[1] in ('ENTER', 'EXIT'):
                    button_key = event[0]
                    if event[1] in 'ENTER':
                        window[button_key].update(text_color='#00D84A',
                                                background_color='#FFF'
                                                )
                    if event[1] in 'EXIT':
                        window[button_key].update(text_color='#009678',
                                                background_color='#FFF'
                                                )
                else:
                    if edit_name is False:
                        window['-CATEGORY_ALIAS-'].update(disabled = False, visible = True)
                        edit_name = True
                    else:
                        window['-CATEGORY_ALIAS-'].update(disabled = True, visible = False)
                        edit_name = False
            if event == '-SAVE-':
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
                try:
                    save_alias(category, str(values['-CATEGORY_ALIAS-']))
                    sg.popup('Mappings were successfully saved.',
                             title='',
                             font = ('Open Sans', 12))
                except KeyError:
                    pass
            if event == '-UNASSIGN_MENU-':
                unassign_menu(page, menu)
                sg.popup(f'The category on...\n\nPage: {page}\nMenu: {menu}\n\nWas successfully unassigned.',
                         title='',
                         font = ('Open Sans', 12))
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 12)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 12)
                )
    window.close()

def main_layout():
    '''Creates the layout for main_window'''
    sg.theme(THEME)
    layout = [
        [sg.Menu(
            MENU_BAR,
            font = ('Open Sans', 10)
            )],
        [sg.Image(rf'{PROGRAM_LOGO}', background_color = '#FFF', pad = (5, (5, 0)))],
        [sg.Text('', background_color = '#FFF')],
        [sg.Text('Choose an exported package file:',
                 font = ('Open Sans', 14, 'bold'),
                 text_color = '#009678',
                 pad = (5, (15, 0)),
                 border_width = 0,
                 background_color = '#FFF'
                 )],
        [sg.Text('', background_color = '#E3FFFA', justification = 'r', size = (35, 1)),
         sg.FileBrowse(key='-EXPORTED_PACKAGES-',
                       enable_events = True,
                       pad = (10, 0),
                       button_color = '#000 on #E3FFFA'
                       )],
        [sg.Text('', background_color = '#FFF')],
        [sg.Button('Create Menu',
                   size = (12, 1),
                   enable_events = True,
                   button_color = '#FFF on #009678',
                   key = '-CREATE_MENU-',
                   pad = ((290, 5), 15)
                   )]
        ]
    return layout

def main():
    '''Handle events for the main window'''
    window = None
    while True:
        try:
            if window is None:
                layout = main_layout()
                window = create_window('Dispensary Menu Creator', layout, PROGRAM_ICON, '#FFF')
                event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Exit':
                break
            # Checks for updates, then asks the user if they want to download it
            if event == 'Check for Updates':
                new_version = update_check()
                if new_version is True:
                    answer = sg.popup_yes_no(
                        UPDATE_MSG_YES,
                        title = 'Update Available',
                        keep_on_top = True,
                        font = ('Open Sans', 12)
                    )
                    if answer == 'Yes':
                        webbrowser.open(DOWNLOAD_LINK)
                else:
                    sg.popup(
                        UPDATE_MSG_NO,
                        title = 'No Updates Available',
                        keep_on_top = True,
                        font = ('Open Sans', 12)
                    )
                window.close()
                window = None
            # Load the 'About' message
            if event == 'About':
                new_version = update_check()
                if new_version is True:
                    sg.popup(
                        ABOUT_UPDATE,
                        title = f'{PROGRAM_NAME}  -  New Version Available',
                        keep_on_top = True,
                        font = ('Open Sans', 12)
                    )
                else:
                    sg.popup(
                        ABOUT,
                        title = f'{PROGRAM_NAME}',
                        keep_on_top = True,
                        font = ('Open Sans', 12)
                        )
                window.close()
                window = None
            # Open a web browser to the GitHub page
            if event == 'GitHub Page':
                webbrowser.open(GITHUB_LINK)
                window.close()
                window = None
            if event == '-CREATE_MENU-':
                exported_packages = pathlib.PurePath(values['-EXPORTED_PACKAGES-'])
                current_packages = pd.read_excel(exported_packages, sheet_name = 'All Packages')
                full_menu = build_menu(current_packages)
                np.sort(full_menu['category'].unique())
                save_all(full_menu)
                window.close()
                window = None
                save_location = pathlib.PurePath(MAIN_DIRECTORY, '_schedules')
                sg.popup(f'The menu was successfully saved to: {save_location}', title='', font = ('Open Sans', 12))
            if event == 'Menu Mapping Configuration':
                window.close()
                cell_map_config()
                window = None
            if event == 'Discounted Products':
                window.close()
                discount_config()
                window = None
        except ValueError as v_error:
            sg.popup_error(
                f'The value {v_error} was out of bounds.',
                title = 'Value Error',
                font = ('Open Sans', 12)
                )
        except TypeError as t_error:
            sg.popup_error(
                f'{t_error} was an improper data type.',
                title = 'Type Error',
                font = ('Open Sans', 12)
                )
    window.close()

if __name__ == '__main__':
    main()
