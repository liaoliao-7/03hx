import pandas as pd


def new_time(rel_time, s, li):  # rel_time，开始时间，str ； s，间隔， int ； li ，个数， int
    nt = []
    t1 = pd.Timestamp(rel_time)
    for i in range(li):
        nt.append(t1)
        t1 += pd.Timedelta(minutes=s)
    return nt
