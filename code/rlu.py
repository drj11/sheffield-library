#!/usr/bin/env python

"""Registered Library Users

Read from ``work/3_10_Registered Library Users.xlsx``
"""

import sys

from collections import OrderedDict

# http://messytables.readthedocs.org/en/latest/
import messytables

FILENAME = "work/3_10_Registered Library Users.xlsx"

def data_rows():
    tableset = messytables.excel.XLSTableSet(filename=FILENAME)
    data = [table for table in tableset.tables if table.name == 'Data']
    assert data, "Failed to find worksheet named 'Data'"
    data, = data
    return data

def data_tuples(rows):
    for row in rows:
        if row[0].value == 'Library':
            continue
        yield tuple((c.value for c in row))

def data_dicts(tuples):
    keys = ['library', 'adult_ru', 'u18_ru', 'total_ru']
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
    sys.exit(main())
