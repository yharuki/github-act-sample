FROM python:3.9-slim

# ENV http_proxy http://192.168.1.13:8080/
# ENV https_proxy http://192.168.1.13:8080/
# ENV http_proxy http://proxy3.nanao.co.jp:8080/
# ENV https_proxy http://proxy3.nanao.co.jp:8080/
# ENV no_proxy localhost,127.0.0.1

ENV WORKSPACE workspace
ENV USERNAME appuser

RUN pip install \
    altgraph==0.17.4 \
    certifi==2024.7.4 \
    charset-normalizer==3.3.2 \
    docopt==0.6.2 \
    et-xmlfile==1.1.0 \
    future==1.0.0 \
    idna==3.8 \
    numpy==1.26.4 \
    openpyxl==3.1.5 \
    pandas==2.2.2 \
    pefile==2024.8.26 \
    pyinstaller==6.10.0 \
    pyinstaller-hooks-contrib==2024.8 \
    python-dateutil==2.9.0.post0 \
    pytz==2024.1 \
    pywin32-ctypes==0.2.3 \
    PyYAML==6.0.2 \
    requests==2.32.3 \
    six==1.16.0 \
    urllib3==2.2.2

RUN mkdir -p /app/tmp
COPY ./app/ /app/

RUN groupadd -g 1000 ${USERNAME}
RUN useradd -m -u 1000 -g 1000 -d /home/${USERNAME} ${USERNAME}
RUN chown -R ${USERNAME}:${USERNAME} /app
#RUN chmod +x /${WORKSPACE}/startup.sh

USER ${USERNAME}
WORKDIR /app

CMD ["python3", "main.py"]