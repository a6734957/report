#!/usr/bin/env python
import os

cmd = os.popen('df -hT').read().strip().split('\n')
for line in cmd:
    line_list = line.split()
    if line_list[6] == '/':
        root = line_list[5].replace("%","")
        root = int(root)
        if root > 80:
            print  "waring !!!! overload %s%%" % root
        else:
             print "used %s%% good !!" % root
