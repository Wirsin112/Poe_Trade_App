import os
import copy
import json
import sqlite3
import requests
import warnings
from math import ceil
from datetime import datetime

connection = sqlite3.connect('item_database.db')
cursor = connection.cursor()

WORST_LIQUIDITY_IN_DAYS = 5


def add_currency_filter_to_query(query, currency):
    """
    Adds the currency filter to the query

    :param query: the query dictionary
    :type query: dict
    :param currency: currency type
    :type currency: str
    :return: the query dictionary
    :rtype: dict
    """
    if "filters" not in query["query"]:
        query["query"]["filters"] = {}
    if "trade_filters" not in query["query"]["filters"]:
        query["query"]["filters"]["trade_filters"] = {}
    if "filters" not in query["query"]["filters"]["trade_filters"]:
        query["query"]["filters"]["trade_filters"]["filters"] = {}
    if "price" not in query["query"]["filters"]["trade_filters"]["filters"]:
        query["query"]["filters"]["trade_filters"]["filters"]["price"] = {}
    if "option" not in query["query"]["filters"]["trade_filters"]["filters"]["price"] or \
            query["query"]["filters"]["trade_filters"]["filters"]["price"]["option"] != currency:
        query["query"]["filters"]["trade_filters"]["filters"]["price"]["option"] = currency

    return query


class Item:

    def __init__(self, name, league, price=None, search_id=None, liquidity=None, date_checked=None, category='item'):
        """
        Class representing an item and its economic activities (data is kept in the database, updated from official trade api)

        :param name: name of the item
        :type name: str
        :param league: name of the league
        :type league: str
        :param price: price of the item (chaos orbs)
        :type price: int
        :param search_id: unique search id
        :type search_id: str
        :param liquidity: measures liquidity (0 - bad, 5 - best)
        :type liquidity: int
        :param date_checked: last time this data was updated
        :type date_checked: datetime
        :param category: 'item' or 'currency'
        :type category: str
        """
        self.name = name
        self.league = league

        self.price = price
        self.search_id = search_id
        self.liquidity = liquidity
        self.date_checked = date_checked
        self.category = category

    @property
    def search_link(self):
        """
        Generates a link to the search website

        :return: link to the item search
        :rtype: str
        """
        if self.category == 'item':
            return f"https://www.pathofexile.com/trade/search/{self.league}/{self.search_id}"
        else:
            return f"https://www.pathofexile.com/trade/exchange/{self.league}/{self.search_id}"

    def get_data_from_api(self):
        """
        Fill all values of the item using a query from 'search_queries'
        The name of the query file in JSON must match the class self.name attribute
        """
        # Load the json query
        query_path = f"search_queries/{self.name}.json"
        if not os.path.exists(query_path):
            raise EnvironmentError(f"There is no \"{query_path}]\" file\nPlease create the query file for {self.name}\n")

        with open(query_path, 'r', encoding='utf8') as f:
            query = json.load(f)

        # Create first request to get list of items that fit the query
        headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

        if self.category == 'item':
            # Process item trades
            url = f"https://www.pathofexile.com/api/trade/search/{self.league}"

            # Request in chaos orbs
            chaos_query = add_currency_filter_to_query(copy.deepcopy(query), "chaos")
            chaos_request = json.loads(requests.post(url, json=chaos_query, headers=headers).text)
            if 'result' not in chaos_request:
                warnings.warn(f"The query for {self.name} (in chaos) returned an invalid response when executed\nCheck if the query is valid",
                              category=RuntimeWarning)
                chaos_result = {}
            else:
                num_chaos_trades = min(10, len(chaos_request['result']))
                chaos_items = ','.join(chaos_request['result'][:num_chaos_trades])
                chaos_id = chaos_request['id']
                chaos_result_url = f"https://www.pathofexile.com/api/trade/fetch/{chaos_items}?query={chaos_id}"
                chaos_result = json.loads(requests.get(chaos_result_url, headers=headers).text)

            # Request in exalted orbs
            exalted_query = add_currency_filter_to_query(copy.deepcopy(query), "exalted")
            exalted_request = json.loads(requests.post(url, json=exalted_query, headers=headers).text)
            if 'result' not in exalted_request:
                warnings.warn(f"The query for {self.name} (in exalted) returned an invalid response when executed\nCheck if the query is valid",
                              category=RuntimeWarning)
                exalted_result = {}
            else:
                num_exalted_trades = min(10, len(exalted_request['result']))
                exalted_items = ','.join(exalted_request['result'][:num_exalted_trades])
                exalted_id = exalted_request['id']
                exalted_result_url = f"https://www.pathofexile.com/api/trade/fetch/{exalted_items}?query={exalted_id}"
                exalted_result = json.loads(requests.get(exalted_result_url, headers=headers).text)

            # Calculate the price and liquidity
            # Chaos
            chaos_price = 0
            chaos_liquidity = 0
            if 'result' in chaos_result:
                chaos_prices = []
                chaos_times = []
                for chaos_offer in chaos_result['result']:
                    chaos_prices.append(chaos_offer["listing"]["price"]["amount"])
                    time_from_now = datetime.utcnow() - datetime.strptime(chaos_offer["listing"]["indexed"], "%Y-%m-%dT%H:%M:%SZ")
                    chaos_times.append(int(time_from_now.total_seconds() / 60))

                # Chaos price = avg(avg, median)
                mean = sum(chaos_prices) / len(chaos_prices)
                median = chaos_prices[int(len(chaos_prices) / 2.0)]
                chaos_price = (mean + median) / 2.0
                # Chaos time = >24h -> 0 | <24h -> up to 5 (inclusive)
                chaos_time_median = sorted(chaos_times)[int(len(chaos_times) / 2.0)]
                chaos_time = (int(sum(chaos_times) / len(chaos_times)) + chaos_time_median) / 2.0
                chaos_liquidity = 5 - min(5, int(chaos_time / (WORST_LIQUIDITY_IN_DAYS * 24 * 60 / 5)))
            # Exalted
            exalted_price = 0
            exalted_liquidity = 0
            if 'result' in exalted_result:
                exalted_prices = []
                exalted_times = []
                for exalted_offer in exalted_result['result']:
                    with connection:
                        cursor.execute("SELECT price FROM items WHERE name='chaos_in_exalt'")
                    exalted_rate = cursor.fetchone()[0]

                    exalted_prices.append(exalted_offer["listing"]["price"]["amount"] * exalted_rate)
                    time_from_now = datetime.utcnow() - datetime.strptime(exalted_offer["listing"]["indexed"], "%Y-%m-%dT%H:%M:%SZ")
                    exalted_times.append(int(time_from_now.total_seconds() / 60))

                # Exalted price = avg(avg, median)
                mean = sum(exalted_prices) / len(exalted_prices)
                median = exalted_prices[int(len(exalted_prices) / 2.0)]
                exalted_price = (mean + median) / 2.0
                # Exalted time = >24h -> 0 | <24h -> up to 5 (inclusive)
                exalted_time_median = sorted(exalted_times)[int(len(exalted_times) / 2.0)]
                exalted_time = (int(sum(exalted_times) / len(exalted_times)) + exalted_time_median) / 2.0
                exalted_liquidity = 5 - min(5, int(exalted_time / (WORST_LIQUIDITY_IN_DAYS * 24 * 60 / 5)))

            if chaos_price == 0:
                chaos_price = exalted_price
            chaos_price = ceil(chaos_price)
            if exalted_price == 0:
                exalted_price = chaos_price
            exalted_price = ceil(chaos_price)

            if chaos_liquidity == 0:
                chaos_liquidity = exalted_liquidity
            if exalted_liquidity == 0:
                exalted_liquidity = chaos_liquidity

            self.price = min(chaos_price, exalted_price)
            self.liquidity = max(chaos_liquidity, exalted_liquidity)

            # Original request for search link generation
            request = json.loads(requests.post(url, json=query, headers=headers).text)
        else:
            # Process currency trades
            url = f"https://www.pathofexile.com/api/trade/exchange/{self.league}"
            request = json.loads(requests.post(url, json=query, headers=headers).text)
            if 'result' not in request:
                warnings.warn(f"The query for {self.name} returned an invalid response when executed\nCheck if the query is valid", category=RuntimeWarning)
                self.price = 0
                self.liquidity = 0
            else:
                num_trades = min(10, len(request['result']))
                items = ','.join(request['result'][:num_trades])
                result_url = f"https://www.pathofexile.com/api/trade/fetch/{items}?query={request['id']}"
                result = json.loads(requests.get(result_url, headers=headers).text)

                if 'result' in result:
                    prices = []
                    for offer in result['result']:
                        prices.append(offer["listing"]["price"]["amount"])
                    self.price = int(prices[-1])
                else:
                    self.price = 100

                self.liquidity = 5

        if 'result' not in request:
            warnings.warn(f"The query for {self.name} returned an invalid response when executed\nCheck if the query is valid", category=RuntimeWarning)
            self.search_id = "Error"
            self.date_checked = datetime.utcnow()
        else:
            self.search_id = request['id']
            self.date_checked = datetime.utcnow()

    def dump_to_database(self):
        """
        Dumps the data to the database
        Either updates the data if item is already present or inserts it
        """
        with connection:
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='items'")

        # Create table if doesn't exist
        if not len(cursor.fetchall()):
            cursor.execute("""CREATE TABLE items (name text,
                                                  league text,
                                                  price integer, 
                                                  search_id text,
                                                  liquidity integer,
                                                  date_checked text,
                                                  category text)""")

        item_dict = {'name': self.name,
                     'league': self.league,
                     'price': self.price,
                     'search_id': self.search_id,
                     'liquidity': self.liquidity,
                     'date_checked': self.date_checked.strftime("%Y-%m-%dT%H:%M:%SZ"),
                     'category': self.category}

        with connection:
            cursor.execute("SELECT * FROM items WHERE name=:name AND league=:league", item_dict)

        if len(cursor.fetchall()):
            # Update if item in database
            if self.price == 0:
                print(f"{self.name:<55} wasn't updated in the database -- Reason: new data was invalid")
            else:
                with connection:
                    cursor.execute("""UPDATE items SET name=:name, 
                                                       league=:league, 
                                                       price=:price, 
                                                       search_id=:search_id, 
                                                       liquidity=:liquidity, 
                                                       date_checked=:date_checked, 
                                                       category=:category
                                                       WHERE name=:name AND league=:league""", item_dict)
                print(f"{self.name:<55} was updated in the database")
        else:
            # Insert if item not in database
            with connection:
                cursor.execute("INSERT INTO items VALUES (:name, :league, :price, :search_id, :liquidity, :date_checked, :category)",
                               item_dict)
            print(f"{self.name:<55} was added to the database")

    def load_from_database(self):
        """
        Loads the item from the database
        """
        with connection:
            cursor.execute("SELECT * FROM items WHERE name=:name AND league=:league", {'name': self.name, 'league': self.league})

        select_result = cursor.fetchall()
        if select_result:
            self.league = select_result[0][1]
            self.price = select_result[0][2]
            self.search_id = select_result[0][3]
            self.liquidity = select_result[0][4]
            self.date_checked = datetime.strptime(select_result[0][5], "%Y-%m-%dT%H:%M:%SZ")
            self.category = select_result[0][6]
        else:
            warnings.warn(f"{self.name} is not in the database", UserWarning)
