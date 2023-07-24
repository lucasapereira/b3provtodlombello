FROM python:3.11-slim


RUN pip install --upgrade pip
RUN pip install numpy
RUN pip install pandas
RUN pip install openpyxl

COPY . .

CMD [ "python", "./application.py" ]