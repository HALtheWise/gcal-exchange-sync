FROM ubuntu:latest

RUN apt-get -y update && \
	apt-get -y install \
	python2.7 \
	git \
	nano \
	cron \
	&& rm -rf /var/lib/apt/lists/*

# Using pip, install pipenv
RUN apt-get -y update && \
	apt-get -y install python-pip && \
	rm -rf /var/lib/apt/lists/* && \
	pip install --upgrade pipenv && \
	apt-get -y autoremove

COPY Pipfile /Pipfile
COPY Pipfile.lock /Pipfile.lock

RUN pipenv install --system

COPY .credentials /.credentials
COPY src /src

# Setup Cron, taken from https://www.ekito.fr/people/run-a-cron-job-with-docker/
COPY crontab /etc/cron.d/calsync-cron
RUN chmod 0644 /etc/cron.d/calsync-cron

CMD ["cron", "-f"]
