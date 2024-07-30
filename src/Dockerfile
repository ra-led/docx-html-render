FROM python:3.9-bullseye

RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y libgdiplus
RUN python3 -m pip install --upgrade pip

WORKDIR /src
COPY requirements.txt requirements.txt
RUN --mount=type=cache,target=/root/.cache/pip pip install -r requirements.txt

# Models load
RUN pip install gdown
RUN python3 -m gdown https://drive.google.com/uc?id=1vuVVl11N93AGKtCW7KxXxFcBh60GmCpF
RUN apt install unzip
RUN unzip model_dir.zip

# Documentation
# RUN pip install sphinx sphinx-autodoc
# RUN make -C docs html

COPY . .
