FROM python:3.11.1-buster

ARG work_dir=/app/

WORKDIR $work_dir

ADD ./requirement.txt $work_dir

RUN apt-get update
RUN apt-get -y install locales && \
    localedef -f UTF-8 -i ja_JP ja_JP.UTF-8
ENV LANG ja_JP.UTF-8
ENV LANGUAGE ja_JP:ja
ENV LC_ALL ja_JP.UTF-8
ENV TZ JST-9
ENV TERM xterm

RUN pip --no-cache-dir install --upgrade pip
RUN pip --no-cache-dir install -r requirement.txt

CMD ["python", "main.py"]
