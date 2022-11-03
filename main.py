import requests
from bs4 import BeautifulSoup
import csv
import pandas as pd
import os


def data_extraction(url):
    res = requests.get(url)
    soup = BeautifulSoup(res.text, 'lxml')
    data = {}

    for item in soup.find_all('h3', class_='header'):
        name = item.find('a').get_text()
        link = item.find('a').get('href')
        # to extract email addresses.
        inner_res = requests.get(link)
        inner_soup = BeautifulSoup(inner_res.text, 'lxml')
        try:
            result = inner_soup.find_all('a', {'title': 'Napísať e-mail'})
            # if person has more than one email
            if len(result) > 1:
                email = ""
                for var in result:
                    email += var.get_text()
            else:
                email = result[0].get_text()
        except (AttributeError, IndexError):
            email = ''
        data[name] = email.replace(" ", ", ")
    return data


def write_into_csv(dictionary):
    with open('result.csv', 'a') as file:
        writer = csv.writer(file)
        for items in dictionary.items():
            writer.writerow(items)


if __name__ == '__main__':
    print("[*] Starting data extraction...")
    url = 'https://obcan.justice.sk/infosud-registre?p_p_id=isufrontreg_WAR_isufront&\
p_p_col_id=column-1&p_p_col_count=1&p_p_mode=view&_isufrontreg_WAR_isufront_entityType=znalec&p_p_state=normal&_isufrontreg_WAR_isufront_view=list&f.23785=25588&_isufrontreg_WAR_isufront_cur='
    # start from page 1
    page = 1
    while True:
        print(f"[*] Now on page: [ {page} ]")
        dictionary = data_extraction(url+str(page))
        if not dictionary:
            print('[*] Nothing to extract, no more pages left')
            break
        print(f'[*] Writing page [ {page} ] into CSV file')
        write_into_csv(dictionary)
        page += 1

    # convert csv into xlsx
    print("[*] Converting CSV file to Excel...")
    read_file = pd.read_csv('result.csv', header=None)
    read_file.to_excel('result.xlsx', index=None, header=False)
    print("[*] Convertion completed... [ result.xslx ] \nfile has been saved into current directory")

    answer = input(
        '[?] Do you want to delete created CSV file? [ type yes for deletion ]: ')
    if answer == 'yes':
        os.remove('result.csv')
        print("[*] Removed CSV file.")
