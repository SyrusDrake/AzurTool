import requests
from bs4 import BeautifulSoup
import pandas as pd
from typing import Any, List
from os.path import exists

url = "https://azurlane.koumakan.jp/wiki/List_of_Ships_by_Stats"


def dl_list() -> pd.DataFrame:
    """Download info from the AL ship list and turn it into a dataframe.

    Returns:
        pd.DataFrame: A full dataframe of ships.
    """
    response = requests.get(url)
    # Makes sure the website can be reached
    if response.status_code != 200:
        print('URL cannot be reached')
        return

    soup = BeautifulSoup(response.text, 'html.parser')
    # Downloads the data from the tabbed tables. There are seven, one for each ship classification
    # TODO: Defaults to Level 1 stats. Possibly change with "article id=..."
    all_tables: Any = soup.find_all(
        'div', {'class': "tabber"})

    # if only find: return is element tag
    # if find_all: return is element result set
    class_list: List[dict] = []
    # Turns the downloaded tables into dictionary, seperately for each ship classification
    for i in range(0, 7):
        class_table: List[Any] = pd.read_html(str(all_tables[i]))
        class_dict: dict = class_table[0].to_dict(orient='records')
        class_list.append(class_dict)

    full_dict: dict = {}
    # Combines all class dictionaries and renames the unnamed categories
    for ship_class in class_list:
        for ship in ship_class:
            temp_dict: dict = {}
            temp_dict['ID'] = ship['ID']
            temp_dict['Rarity'] = ship['Rarity']
            temp_dict['Got?'] = ''
            temp_dict['Nation'] = ship['Nation']
            temp_dict['Type'] = ship['Type']
            temp_dict['Luck'] = ship['Unnamed: 5']
            temp_dict['Armor'] = ship['Unnamed: 6']
            temp_dict['Speed'] = ship['Spd']
            temp_dict['Health'] = ship['Unnamed: 8']
            temp_dict['Firepower'] = ship['Unnamed: 9']
            temp_dict['AA'] = ship['Unnamed: 10']
            temp_dict['Torpedo'] = ship['Unnamed: 11']
            temp_dict['Evasion'] = ship['Unnamed: 12']
            temp_dict['Aviation'] = ship['Unnamed: 13']
            temp_dict['Oil'] = ship['Unnamed: 14']
            temp_dict['Reload'] = ship['Unnamed: 15']
            temp_dict['ASW'] = ship['Unnamed: 16']
            temp_dict['Oxygen'] = ship['Unnamed: 17']
            temp_dict['Ammo'] = ship['Unnamed: 18']
            temp_dict['Accuracy'] = ship['Unnamed: 19']

            full_dict[ship['Ship Name']] = temp_dict

    df_ships = (pd.DataFrame(data=full_dict)).transpose()
    return df_ships


def compare_list(df_existing: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    """Compares an existing and a new dataframe and adds new items to the existing one.

    Args:
        df_existing (pd.DataFrame): The data frame extracted from the existing excel list.
        df_new (pd.DataFrame): The data frame with newly downloaded data.

    Returns:
        pd.DataFrame: Existing dataframe with new items added.
    """
    df_existing = pd.concat([df_existing, df_new])
    # checkship = df_existing.loc['Vestal']
    # print(type(checkship.iloc[0]['ID']))
    # print(type(checkship.iloc[1]['ID']))
    df_existing.drop_duplicates(subset='ID', keep='first', inplace=True)
    return df_existing


def create_excel(dataframe: pd.DataFrame) -> None:
    """Creates an excel file from the dataframe, formats and saves it.

    Args:
        dataframe (pd.DataFrame): The dataframe from which the file should be created.
    """

    # Drops potential empty rows
    dataframe.dropna(subset=['ID'], inplace=True)

    # Replaces all Collab "nations" with a common descriptor
    nations = ['Royal Navy', 'Eagle Union', 'Sakura Empire', 'Iron Blood', 'Dragon Empery',
               'Northern Parliament', 'Iris Libre', 'Vichya Dominion', 'Sardegna Empire', 'META', 'Universal']
    dataframe.loc[~dataframe["Nation"].isin(nations), "Nation"] = "Collab"

    writer = pd.ExcelWriter('ships.xlsx', engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    (max_row, max_col) = dataframe.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    # Sets ID column to text to preserve leading zeros
    id_format = workbook.add_format({'num_format': '@'})
    worksheet.set_column(1, 1, None, id_format)

    # Colours all ships based on rarity
    normal = workbook.add_format({'bg_color':   '#d5d5d5',
                                  'font_color': '#000000'})
    rare = workbook.add_format({'bg_color':   '#baefff',
                                'font_color': '#000000'})
    elite = workbook.add_format({'bg_color':   '#d6c7ff',
                                'font_color': '#000000'})
    superrare = workbook.add_format({'bg_color':   '#eeeecd',
                                    'font_color': '#000000'})
    ultrarare = workbook.add_format({'bg_color':   '#b5ffd8',
                                    'font_color': '#000000'})

    worksheet.conditional_format(0, 0, max_row, 2,
                                 {'type':     'formula',
                                  'criteria': '=$C1="Normal"',
                                  'format':    normal})
    worksheet.conditional_format(0, 0, max_row, 2,
                                 {'type':     'formula',
                                  'criteria': '=$C1="Rare"',
                                  'format':    rare})
    worksheet.conditional_format(0, 0, max_row, 2,
                                 {'type':     'formula',
                                  'criteria': '=$C1="Elite"',
                                  'format':    elite})
    worksheet.conditional_format(0, 0, max_row, 2,
                                 {'type':     'formula',
                                  'criteria': '=$C1="Priority"',
                                  'format':    superrare})
    worksheet.conditional_format(0, 0, max_row, 2,
                                 {'type':     'formula',
                                  'criteria': '=$C1="Super Rare"',
                                  'format':    superrare})
    worksheet.conditional_format(0, 0, max_row, 2,
                                 {'type':     'formula',
                                  'criteria': '=$C1="Ultra Rare"',
                                  'format':    ultrarare})
    worksheet.conditional_format(0, 0, max_row, 2,
                                 {'type':     'formula',
                                  'criteria': '=$C1="Decisive"',
                                  'format':    ultrarare})
    writer.save()


def main():
    df_new = dl_list()

    if exists('ships.xlsx'):
        with open('ships.xlsx', 'rb') as input_file:
            df_existing = pd.read_excel(input_file)
            df_existing.set_index(df_existing.columns[0], inplace=True)

        df_existing = compare_list(df_existing, df_new)
        create_excel(df_existing)
    else:
        create_excel(df_new)


if __name__ == '__main__':
    main()
