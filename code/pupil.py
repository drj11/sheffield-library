#!/usr/bin/env python

"""Low Attaining Pupils
"""

import sys

from collections import OrderedDict

# http://messytables.readthedocs.org/en/latest/
import messytables

FILENAME = "work/3_8_Low attaining pupils.xls"

def data_rows():
    tableset = messytables.excel.XLSTableSet(filename=FILENAME)
    data = [table for table in tableset.tables if table.name == 'Data']
    assert data, "Failed to find worksheet named 'Data'"
    data, = data
    return data

def data_tuples(rows):
    rows = iter(rows)
    for row in rows:
        if row[0].value.startswith('Library branch'):
            break
    for row in rows:
        yield tuple((c.value for c in row))

def data_dicts(tuples):
    keys = ['library',
      'foundation_low_percent', 'foundation_total', 'foundation_low',
      'ks1_low_percent', 'ks1_total', 'ks1_low',
      'ks2_low_percent', 'ks2_total', 'ks2_low',
      'ks4_low_percent', 'ks4_total', 'ks4_low',
      'total_low',
      'rank_low', 'rank_percent', 'rank_overall'
      ]
    for t in tuples:
        yield OrderedDict(zip(keys, t))

def data():
    """
    Return the Registered Library User data as a sequence of
    (Ordered) dicts, one per library, where each dict has the
    keys = ['library',
      'foundation_low_percent', 'foundation_total', 'foundation_low',
      'ks1_low_percent', 'ks1_total', 'ks1_low',
      'ks2_low_percent', 'ks2_total', 'ks2_low',
      'ks4_low_percent', 'ks4_total', 'ks4_low',
      'total_low',
      'rank_low', 'rank_percent', 'rank_overall'
      ]
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
