FROM amazoncorretto:21

# Install LibreOffice and fonts
RUN yum install -y \
    libXinerama \
    cups-libs \
    cairo \
    wget \
    tar \
    && cd /tmp \
    && wget https://ftp.osuosl.org/pub/tdf/libreoffice/stable/25.8.4/rpm/x86_64/LibreOffice_25.8.4_Linux_x86-64_rpm.tar.gz \
    && tar -xzf LibreOffice_25.8.4_Linux_x86-64_rpm.tar.gz \
    && cd LibreOffice_25.8.4*_rpm/RPMS \
    && yum localinstall -y *.rpm \
    && ln -sf /opt/libreoffice25.8/program/soffice /usr/local/bin/libreoffice \
    && rm -rf /tmp/*

# Install MS P Gothic compatible font
RUN yum install -y google-noto-sans-cjk-fonts

COPY target/pdf-converter-s3-1.0.0.jar /app/app.jar
WORKDIR /app
CMD ["java", "-jar", "app.jar"]