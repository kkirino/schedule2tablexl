#!/bin/bash
### This shellscript can run Python GUI App on WSL2 ###
docker run --rm --env DISPLAY=`cat /etc/resolv.conf | grep nameserver | awk '{print $2}'`:0.0 \
--volume `pwd`:/src --entrypoint /bin/sh python_gui_app -c "python main.py"
