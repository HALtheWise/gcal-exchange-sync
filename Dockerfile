FROM ubuntu:latest

# Following instructions from https://github.com/google/cloud-print-connector/wiki/Build-from-source

RUN apt-get -y update && \
	apt-get -y install \
	python2.7 \
	git \
	nano \
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
CMD ["python2.7","/src/Synchronizer.py"]
#COPY
