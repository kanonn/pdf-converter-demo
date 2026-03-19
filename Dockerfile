FROM amazonlinux:2023

ENV LANG=en_US.UTF-8
ENV LC_ALL=en_US.UTF-8

RUN dnf install -y \
    java-21-amazon-corretto-devel \
    glibc-langpack-en \
    libXinerama \
    cups-libs \
    cairo \
    nss \
    # --- ---
    libXrender \
    libXext \
    libSM \
    libICE \
    fontconfig \
    # --- ---
    libX11 \
    # ---  ---
    libXau \
    libXcomposite \
    libXdamage \
    libXfont2 \
    libXi \
    libXrandr \
    libXt \
    libxcb \
    dbus \
    at-spi2-atk \
    atk \
    gdk-pixbuf2 \
    pango \
    cairo-gobject \
    xcb-util \
    # -----------------------------------------------
    openssl \
    openssl-libs \
    tar \
    gzip \
    google-noto-sans-cjk-fonts && \
    dnf clean all  # 架构师提醒：这一层已经 1.28G 了，必须保持 clean，防止体积失控

# Copy pre-downloaded LibreOffice from build context
COPY LibreOffice_25.8.5_Linux_x86-64_rpm.tar.gz /tmp/

# Install LibreOffice
RUN cd /tmp \
    && tar -xzf LibreOffice_25.8.5_Linux_x86-64_rpm.tar.gz \
    && cd LibreOffice_25.8.5*_rpm/RPMS \
    && dnf install -y *.rpm \
    && rm -rf /tmp/*

# Create symlink (check actual path after install)
RUN ln -sf /opt/libreoffice25.8/program/soffice /usr/local/bin/libreoffice \
    || ln -sf /opt/libreoffice25.8.5/program/soffice /usr/local/bin/libreoffice \
    || echo "Symlink creation - checking available paths" && ls /opt/

# Verify
RUN /usr/local/bin/libreoffice --version || ls -la /opt/libreoffice*/program/

COPY target/pdf-converter-s3-1.0.0.jar /app/app.jar
WORKDIR /app

CMD ["java", "-Dfile.encoding=UTF-8", "-jar", "app.jar"]