import os
import json
import time
import urllib3
import requests
import warnings
import xlsxwriter
from random import shuffle
from item import Item
from generate_excel import generate_excel

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
league = json.loads(requests.get("https://www.pathofexile.com/api/trade/data/leagues", headers=headers).text)['result'][0]['id']


def main():
    # Refresh the chaos-exalt query
    item = Item(name='chaos_in_exalt', league=league, category='currency')
    item.get_data_from_api()
    item.dump_to_database()
    time.sleep(30)

    # Refresh all other queries (random order to improve consistency of the data)
    list_of_query_files = os.listdir("search_queries")
    shuffle(list_of_query_files)
    for search_query_filename in list_of_query_files:
        query_name = os.path.basename(search_query_filename).rsplit('.', 1)[0]
        if query_name == 'chaos_in_exalt':
            item = Item(name=query_name, league=league, category='currency')
        else:
            item = Item(name=query_name, league=league, category='item')
        item.get_data_from_api()
        item.dump_to_database()

        # Try updating excel file
        try:
            generate_excel()
        except (xlsxwriter.exceptions.FileCreateError, TypeError):
            print('Cannot update excel - close the workbook or wait until all items are downloaded!')
        time.sleep(60)


if __name__ == "__main__":
    while True:
        try:
            main()
        except urllib3.exceptions.NewConnectionError:
            warnings.warn("Could not connect to the Path Of Exile API, please check your connection!", category=RuntimeWarning)
            time.sleep(60)
