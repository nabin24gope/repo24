FROM python:3.8

WORKDIR /app

COPY ./ Event.py

RUN pip install pandas numpy requests openpyxl random datetime xlsxwriter glob shutil 


CMD ["python","./Event.py"]


