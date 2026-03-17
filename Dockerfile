# ============================================================
# Stage 1: Build
# ============================================================
FROM amazoncorretto:21 AS builder

WORKDIR /build
COPY pom.xml .
RUN mvn dependency:go-offline -B 2>/dev/null || true

COPY src ./src
RUN mvn clean package -DskipTests -B

# ============================================================
# Stage 2: Runtime (Red Hat UBI 9)
# ============================================================
FROM redhat/ubi9

# --- Java 21 (Amazon Corretto) ---
RUN dnf install -y wget tar fontconfig && \
    # Corretto
    rpm --import https://yum.corretto.aws/corretto.key && \
    curl -L -o /etc/yum.repos.d/corretto.repo \
        https://yum.corretto.aws/corretto.repo && \
    dnf install -y java-21-amazon-corretto-headless && \
    dnf clean all

# --- LibreOffice ---
RUN cd /tmp && \
    wget -q https://download.documentfoundation.org/libreoffice/stable/24.2.7/rpm/x86_64/LibreOffice_24.2.7_Linux_x86-64_rpm.tar.gz && \
    tar -xzf LibreOffice_24.2.7_Linux_x86-64_rpm.tar.gz && \
    cd LibreOffice_24.2.7.2_Linux_x86-64_rpm/RPMS && \
    dnf localinstall -y *.rpm && \
    cd /tmp && rm -rf LibreOffice_* && \
    ln -sf /opt/libreoffice24.2/program/soffice /usr/local/bin/libreoffice && \
    dnf clean all

# --- CJK fonts ---
RUN dnf install -y \
    google-noto-sans-cjk-ttc-fonts \
    google-noto-serif-cjk-ttc-fonts && \
    fc-cache -fv && \
    dnf clean all

# --- LibreOffice runtime dependencies ---
RUN dnf install -y \
    libX11 \
    libXext \
    libXrender \
    mesa-libGL \
    && dnf clean all

# --- App setup ---
WORKDIR /app
COPY --from=builder /build/target/pdf-converter-s3-*.jar app.jar

RUN mkdir -p /app/temp /app/output

ENV JAVA_HOME=/usr/lib/jvm/java-21-amazon-corretto
ENV AWS_REGION=ap-northeast-1
ENV HOME=/tmp

CMD ["java", "-Djava.io.tmpdir=/tmp", "-jar", "app.jar"]