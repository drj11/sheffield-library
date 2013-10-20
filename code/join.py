#!/usr/bin/env python

from collections import defaultdict

import issuegbp
import pupil
import rluper
import rlu
import visit

def data():
    res = defaultdict(dict)
    for source in [issuegbp, pupil, rluper, rlu, visit]:
        for d in source.data():
          res[d['library']].update(d)
    return dict(res)

def main():
    for d in sorted(data().iteritems()):
        print d

if __name__ == '__main__':
    main()
