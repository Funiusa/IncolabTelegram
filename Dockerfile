FROM python:3.10 as build

EXPOSE 80

WORKDIR /app
ADD ./Bot/. ./
COPY ./requirements.txt ./

RUN pip install -r requirements.txt

CMD ["python", "bot_main.py"]

