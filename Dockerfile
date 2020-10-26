FROM alpine:latest

RUN apk update \
    && apk add --no-cache \
        python3 \
        gcc \
        linux-headers \
        python3-dev \
        libxml2-dev \
        libxslt-dev \
        libffi-dev \
        musl-dev \
        openssl-dev \
        curl

RUN curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py \
    && python3 get-pip.py

RUN pip install \
        exchangelib \
        click \
        PyYaml \
        slugify

COPY . /ews

ENTRYPOINT ["python3", "/ews/thumbscr-ews.py"]