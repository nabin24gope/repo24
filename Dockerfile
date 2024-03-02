FROM python:3.8.18-bullseye

WORKDIR /app

RUN /usr/local/bin/python -m pip install --upgrade pip

RUN pip install pandas 
RUN pip intsall numpy
RUN pip intsall requests
RUN pip install openpyxl
RUN pip install random
RUN pip install datetime
RUN pip install xlsxwriter
RUN pip install glob
RUN pip install shutil 

COPY ./ Event.py
CMD ["python","./Event.py"]


