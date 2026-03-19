FROM amazonlinux:2023

# Set locale
ENV LANG=en_US.UTF-8
ENV LC_ALL=en_US.UTF-8

# Install dependencies
RUN dnf install -y \
    java-21-amazon-corretto-devel \
    glibc-langpack-en \
    libXinerama \
    cups-libs \
    cairo \
    openssl \
    openssl-libs \
    wget \
    tar \
    gzip \
    google-noto-sans-cjk-fonts

# Install LibreOffice
RUN cd /tmp \
    && wget https://ftp.osuosl.org/pub/tdf/libreoffice/stable/25.8.4/rpm/x86_64/LibreOffice_25.8.4_Linux_x86-64_rpm.tar.gz \
    && tar -xzf LibreOffice_25.8.4_Linux_x86-64_rpm.tar.gz \
    && cd LibreOffice_25.8.4*_rpm/RPMS \
    && dnf install -y *.rpm \
    && ln -sf /opt/libreoffice25.8/program/soffice /usr/local/bin/libreoffice \
    && rm -rf /tmp/*

# Verify installation
RUN libreoffice --version

COPY target/pdf-converter-s3-1.0.0.jar /app/app.jar
WORKDIR /app

CMD ["java", "-Dfile.encoding=UTF-8", "-jar", "app.jar"]