import random
import numpy as np


def deal(st, b=None, te=None, tf=None):
    alt = np.array(st)
    var = True

    for i in alt:
        if i < 30 or i > 33:
            tap = list(np.where(alt == i))
            alt[int(tap[0][0])] = round(random.uniform(31, 33), 2)

    while var:
        if alt.std(ddof=1) / alt.mean() <= 0.02:
            b = np.around(alt.std(ddof=1) / alt.mean(), 4)
            var = False
        else:
            temp1 = np.max(alt) - 0.4
            temp2 = np.min(alt) + 0.4
            if len(np.where(alt == np.max(alt))[0]) == 1:
                te = int(np.where(alt == np.max(alt))[0])
            else:
                te = int(np.where(alt == np.max(alt))[0][0])
            if len(np.where(alt == np.min(alt))[0]) == 1:
                tf = int(np.where(alt == np.min(alt))[0])
            else:
                tf = int(np.where(alt == np.min(alt))[0][0])
            alt[te] = temp1
            alt[tf] = temp2

    return alt, b
