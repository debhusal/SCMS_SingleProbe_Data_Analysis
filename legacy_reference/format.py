import sys
import numpy as np


data = np.loadtxt(sys.argv[1])
data_low = data * 0.999995
data_high = data * 1.000005

with open(sys.argv[2], 'w') as f:
    for i in range(len(data)):
        f.write("| (df$m.z>%8.4f & df$m.z<%8.4f)\n" % (data_low[i], data_high[i]))
