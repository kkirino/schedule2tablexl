# FOR DEVELOPMENT USE ONLY
# DO NOT USE THIS CONTAINER FOR DEPOYMENT
FROM cdrx/pyinstaller-windows:python3 
COPY requirements.txt /src
RUN cd /src && pip install -r requirements.txt
WORKDIR /src
