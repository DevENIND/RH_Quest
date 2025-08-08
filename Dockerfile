FROM python:3.13

WORKDIR /app

COPY . .

RUN pip install --no-cache-dir flet

EXPOSE 8000

CMD ["python", "app.py"]