#!/bin/bash
version=0.1.2
docker run --rm --volume `pwd`:/src --entrypoint /bin/sh cdrx/pyinstaller-windows:python3 \
-c "pip install -r requirements.txt && pyinstaller main.py --noconsole --onefile --clean --name schedule2tablexl-${version}"
