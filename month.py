#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import os.path
import sys

from agent import Agent, is_excel
# from refresh import is_excel, refresh

if len(sys.argv) > 1:
    root_dir = sys.argv[1]
else:
    root_dir = os.path.expanduser("~/Downloads/dashifu/")

root_dir = os.path.dirname(root_dir)

if len(sys.argv) > 2:
    output_fn = sys.argv[2]
else:
    output_fn = "./out.txt"


infos = []

for root, _, filenames in os.walk(root_dir):
    for filename in filenames:
        fn = os.path.join(root, filename)
        if not is_excel(fn):
            continue
        print "handling {}".format(fn)
        agent = Agent(fn)
        if not agent.is_valid:
            continue
        if not agent.parse():
            agent.close()
            print "{} failed! error={}".format(fn, agent.error_msg)
            continue
        info = agent.get_month_info(root_dir)
        if len(info) != 0:
            infos.append(info)

max_length = max(len(x) for x in infos)

print "all files done"

for i in xrange(len(infos)):
    l = len(infos[i])
    while l < max_length:
        infos[i].insert(0, "")
        l += 1


with open(output_fn, "w") as fp:
    for info in infos:
        fp.write("\t".join(map(str, info)))
        fp.write("\n")
