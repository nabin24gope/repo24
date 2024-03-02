FROM python:3.8.18-bullseye



RUN /usr/local/bin/python -m pip install --upgrade pip

RUN pip install nbterm numpy matplotlib seaborn pandas


COPY  . .

CMD ["python","./Event.py"]


