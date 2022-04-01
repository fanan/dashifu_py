#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import os.path
import sys

from agent import Agent, is_excel

if len(sys.argv) == 2:
    root_dir = sys.argv[1]
else:
    root_dir = os.path.expanduser("~/Downloads/dashifu/2021")

for root, _, filenames in os.walk(root_dir):
    for filename in filenames:
        fn = os.path.join(root, filename)
        if not is_excel(fn):
            continue
        if not fn.endswith(".xlsx"):
            continue
        print("handling {}".format(fn))
        agent = Agent(fn)
        if not agent.is_valid:
            continue
        if not agent.parse():
            agent.close()
            print("{} failed! error={}".format(fn, agent.error_msg))
            continue
        agent.close()
        new_fn = fn.replace("2021", "2022")
        new_dir = os.path.dirname(new_fn)
        if not os.path.exists(new_dir):
            print("mkdir -p {}".format(new_dir))
            os.makedirs(new_dir)
        try:
            agent.newyear(new_fn)
        except Exception as e:
            print(e)
        # os.rename(fn, new_fn)

