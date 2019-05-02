from datetime import datetime
import json


class Program:

    def __init__(self, initiative, strategy, activity, name):
        self.numYears = 3
        self.firstYear = 2018

        self.initiative = initiative
        self.strategy = strategy
        self.activity = activity
        self.name = name

        self.events = {'totalNumEvents': 0,
                       'numPeopleTotal': 0,
                       'numPeopleUniqueTotal': 0}

        for year in range(self.firstYear, self.firstYear + self.numYears):
            for month in range(1, 13):
                self.events[str(year) + '-' + str(month)] = {
                    'numEvents': 0,
                    'numPeople': 0,
                    'numPeopleUnique': 0
                }

    def addEvent(self, date: datetime, numPeople, numPeopleUnique):
        if numPeople is None:
            numPeople = 0
        if numPeopleUnique is None:
            numPeopleUnique = 0

        self.events['totalNumEvents'] += 1
        self.events['numPeopleTotal'] += numPeople
        self.events['numPeopleUniqueTotal'] += numPeopleUnique

        self.events[str(date.year) + '-' + str(date.month)]['numEvents'] += 1
        self.events[str(date.year) + '-' + str(date.month)]['numPeople'] += numPeople
        self.events[str(date.year) + '-' + str(date.month)]['numPeopleUnique'] += numPeopleUnique

    # returns a list with a dictionary for each month of the year
    def getMonthDict(self, year):
        list = []
        for month in range(1, 13):
            list.append({
                'numEvents': self.events[str(year) + '-' + str(month)]['numEvents'],
                'numPeople': self.events[str(year) + '-' + str(month)]['numPeople'],
                'numPeopleUnique': self.events[str(year) + '-' + str(month)]['numPeopleUnique'],
            })
        return list

    # returns the last month and year in which the program has entries
    def getLastYearAndMonthWithEntries(self):
        for year in reversed(range(self.firstYear, self.firstYear + self.numYears)):
            for month in reversed(range(1, 13)):
                if self.events[str(year) + '-' + str(month)]['numPeople'] > 0:
                    return month, year
        return 0, 0

    def __str__(self):
        return '%s - NumPeople: %s - NumPeopleUnique: %s' % (
            self.name, self.events['numPeopleTotal'], self.events['numPeopleUniqueTotal'])
        # return json.dumps(self.events, indent=4)
