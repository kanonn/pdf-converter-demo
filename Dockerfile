# Use Amazon Linux 2023 (has GLIBC 2.34)
FROM amazonlinux:2023

# Install Java
RUN dnf install -y java-21-amazon-corretto-devel

# Install LibreOffice dependencies
RUN dnf install -y \
    libXinerama \
    cups-libs \
    cairo \
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

COPY target/pdf-converter-s3-1.0.0.jar /app/app.jar
WORKDIR /app
CMD ["java", "-jar", "app.jar"]