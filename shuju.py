import numpy as np


def some_su(a=None, b=None):
    var = True
    while var:
        number = np.random.randint(30, 32, size=20)  # 整数部分
        fu_dian = np.around(np.random.random(20), 4)  # 小数部分
        a = number + fu_dian
        if np.var(a) / np.mean(a) <= 0.02:
            b = np.around(np.var(a) / np.mean(a), 4)
            var = False
    return a, b


