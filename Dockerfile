FROM python:3-alpine
ADD . /nmo
WORKDIR /nmo
# Update the package index
RUN apk update

# Install sudo
RUN apk add --no-cache sudo

# Install SSH client
RUN apk update && apk add openssh-client

# Set environment variables for SSH client
ENV SSH_CLIENT='ssh' \
    SSH_AUTH_SOCK='/ssh-agent' \
    SSH_TTY='/dev/tty'

# Add the following line to specify the HostKeyAlgorithms option
RUN echo 'HostKeyAlgorithms +ssh-dss' >> /etc/ssh/ssh_config

# Install the MySQL client
RUN apk add --no-cache mysql-client

# Install sshpass
RUN apk add --no-cache sshpass

RUN pip install --no-cache-dir markupsafe==2.0.1 Flask sshtunnel mysql-connector-python
RUN pip install -r requirements.txt
EXPOSE 5000

