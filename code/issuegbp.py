#!/usr/bin/env python

"""Issue per GBP
"""

import sys

from collections import OrderedDict

# http://messytables.readthedocs.org/en/latest/
import messytables

FILENAME = "work/3_6_Issues per.xlsx"

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
        if not any(c.value for c in row):
            break
        yield tuple((c.value for c in row))

def data_dicts(tuples):
    keys = ['library', 'issues', 'pn', 'total_issues',
      'budget', 'issues_per_gbp']
    for t in tuples:
        yield OrderedDict(zip(keys, t))

def data():
    """
    Return the Registered Library User data as a sequence of
    (Ordered) dicts, one per library, where each dict has the
    keys = ['library', 'issues', 'pn', 'total_issues',
      'budget', 'issues_per_gbp']
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
