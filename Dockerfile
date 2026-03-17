# ============================================================
# Stage 1: Build (Maven build)
# ============================================================
FROM amazoncorretto:21 AS builder

WORKDIR /build

# Copy Maven files first (for layer caching)
COPY pom.xml .
RUN mvn dependency:go-offline -B 2>/dev/null || true

COPY src ./src
RUN mvn clean package -DskipTests -B

# ============================================================
# Stage 2: Runtime
# ============================================================
FROM amazonlinux:2023

# --- Java 21 ---
RUN dnf install -y java-21-amazon-corretto-headless && \
    dnf clean all

# --- LibreOffice (not in AL2023 standard repo, download RPM directly) ---
RUN dnf install -y wget tar && \
    cd /tmp && \
    wget -q https://download.documentfoundation.org/libreoffice/stable/24.2.7/rpm/x86_64/LibreOffice_24.2.7_Linux_x86-64_rpm.tar.gz && \
    tar -xzf LibreOffice_24.2.7_Linux_x86-64_rpm.tar.gz && \
    cd LibreOffice_24.2.7.2_Linux_x86-64_rpm/RPMS && \
    dnf localinstall -y *.rpm && \
    cd /tmp && rm -rf LibreOffice_* && \
    ln -sf /opt/libreoffice24.2/program/soffice /usr/local/bin/libreoffice && \
    dnf clean all

# --- CJK fonts (Japanese) ---
RUN dnf install -y google-noto-sans-cjk-fonts google-noto-serif-cjk-fonts && \
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

# Copy built JAR from builder stage
COPY --from=builder /build/target/pdf-converter-s3-*.jar app.jar

# Temp directories for LibreOffice
RUN mkdir -p /tmp/libreoffice-profile && \
    mkdir -p /app/temp && \
    mkdir -p /app/output

# Environment variables (override at task definition level)
ENV JAVA_HOME=/usr/lib/jvm/java-21-amazon-corretto
ENV AWS_REGION=ap-northeast-1

# LibreOffice headless profile directory
ENV HOME=/tmp

CMD ["java", \
     "-Djava.io.tmpdir=/tmp", \
     "-jar", "app.jar"]