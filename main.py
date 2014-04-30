#!/usr/bin/python

import os
for root, dirs, files in os.walk("./MOD04files"):
    for file in files:
        if file.endswith(".hdf"):
        	##print(file)
            os.system("python nasa.py "+ os.path.join(root, file));