FROM eclipse-temurin:21-jre-jammy

# --- LibreOffice + CJK fonts ---
RUN apt-get update && \
    apt-get install -y \
        libreoffice-core \
        libreoffice-calc \
        libreoffice-writer \
        fonts-noto-cjk \
    && apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# --- App setup ---
WORKDIR /app
COPY target/pdf-converter-s3-*.jar app.jar

RUN mkdir -p /app/temp /app/output

ENV AWS_REGION=ap-northeast-1
ENV HOME=/tmp

CMD ["java", "-Djava.io.tmpdir=/tmp", "-jar", "app.jar"]