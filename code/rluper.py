#!/usr/bin/env python

"""Registered Library User per Adult"""

import sys

from collections import OrderedDict

# http://messytables.readthedocs.org/en/latest/
import messytables

FILENAME = "work/3_11_RLUs per adult population.xlsx"

def data_rows():
    tableset = messytables.excel.XLSTableSet(filename=FILENAME)
    data = [table for table in tableset.tables
      if table.name == 'Sheet1']
    assert data, "Failed to find worksheet named 'Sheet1'"
    data, = data
    return data

def data_tuples(rows):
    for row in rows:
        if not row[0].value:
            continue
        if row[0].value == "Hillsboro'":
            row[0].value = u'Hillsborough'
        yield tuple((c.value for c in row))

def data_dicts(tuples):
    keys = ['library', 'rlu', 'ge19_population']
    for t in tuples:
        yield OrderedDict(zip(keys, t))

def data():
    """
    Return the Registered Library User data as a sequence of
    (Ordered) dicts, one per library, where each dict has the
    keys: 'library', 'adult_ru', 'u18_ru', 'total_ru'.
    """

    rows = data_rows()
    tuples = data_tuples(rows)
    dicts = data_dicts(tuples)
    return dicts

def main():
    for d in data():
        print d

if __name__ == '__main__':
    main()
