FROM ubuntu:latest

RUN apt-get -y update && \
	apt-get -y install \
	python2.7 \
	git \
	nano \
	cron \
	&& rm -rf /var/lib/apt/lists/*

# Using pip, install the various dependencies
RUN apt-get -y update && \
	apt-get -y install python-pip && \
	rm -rf /var/lib/apt/lists/* && \
	pip install --upgrade pyexchange \
		httplib2 \
		google-api-python-client \
		unidecode \
	&& apt-get -y purge python-pip && \
	apt-get -y autoremove


COPY .credentials /.credentials
COPY src /src

# Setup Cron, taken from https://www.ekito.fr/people/run-a-cron-job-with-docker/
COPY crontab /etc/cron.d/calsync-cron
RUN chmod 0644 /etc/cron.d/calsync-cron

CMD ["cron", "-f"]
