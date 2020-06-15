#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import os.path
import sys

from agent import Agent
from refresh import is_excel, refresh

if len(sys.argv) == 2:
    root_dir = sys.argv[1]
else:
    root_dir = os.path.expanduser("~/Downloads/dashifu/2018")

for root, _, filenames in os.walk(root_dir):
    for filename in filenames:
        fn = os.path.join(root, filename)
        if not is_excel(fn):
            continue
        if not fn.endswith(".xls"):
            continue
        print "handling {}".format(fn)
        try:
            refresh(fn)
        except Exception as e:
            print "{}:{}".format(fn, e.message)
            continue
        else:
            print "refreshed {}".format(fn)
            agent = Agent(fn)
            if not agent.is_valid:
                continue
            if not agent.parse():
                agent.close()
                print "{} failed! error={}".format(fn, agent.error_msg)
                continue
            agent.close()
            agent.newyear(fn)
            new_fn = fn.replace("2018", "2019")
            os.rename(fn, new_fn)

