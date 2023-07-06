FROM python:3.7.2
ADD . /nmo
WORKDIR /nmo
RUN pip install markupsafe==2.0.1
RUN pip install -r requirements.txt
EXPOSE 5000

