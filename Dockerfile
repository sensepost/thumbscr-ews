FROM python as builder

COPY . /ews
WORKDIR /wheels
RUN pip wheel /ews

FROM python:slim

COPY --from=builder /wheels /wheels

RUN pip install -f /wheels thumbscr-ews
RUN rm -rf /wheels

ENTRYPOINT ["thumbscr-ews"]
