from item import Item


class Recipe:

    def __init__(self, name, league, components, results, wiki):
        """
        Recipe object that uses Item objects to calculate profitability

        :param name: name of the recipe
        :type name: str
        :param league: current league name
        :type league: str
        :param components: list of component objects and their counts
        :type components: list[list[str, float]]
        :param results: list of result objects and their counts
        :type results: list[list[str, float]]
        :param wiki: link to the game wiki
        :type wiki: str
        """
        self.name = name
        self.league = league
        self.wiki = wiki

        # Save components, counts and sum up the costs
        self.cost = 0
        self.components = []
        for component in components:
            item = Item(name=component[0], league=self.league)
            item.load_from_database()
            self.cost += item.price * component[1]
            self.components.append([item, component[1]])

        # Save results, counts and sum up the revenue
        self.revenue = 0
        self.results = []
        for result in results:
            item = Item(name=result[0], league=self.league)
            item.load_from_database()
            self.revenue += item.price * result[1]
            self.results.append([item, result[1]])

        self.profit = self.revenue - self.cost
        if self.cost == 0:
            self.roi = 0
        else:
            self.roi = round(100 * self.profit / self.cost, 1)
