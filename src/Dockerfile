FROM python:3.9-bullseye

RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y libgdiplus
RUN python3 -m pip install --upgrade pip

WORKDIR /src
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 5000
COPY . .
CMD ["python", "app.py"]
