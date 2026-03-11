#!/bin/bash
#
# EC2セットアップスクリプト（Amazon Linux 2023対応版）
# 使用方法: chmod +x setup_ec2.sh && ./setup_ec2.sh
#

set -e

echo "=========================================="
echo "  PDF Converter Demo - EC2セットアップ"
echo "=========================================="

# OS判定
if [ -f /etc/os-release ]; then
    . /etc/os-release
    OS=$ID
    VERSION_ID=$VERSION_ID
else
    OS="unknown"
fi

echo "検出されたOS: $OS $VERSION_ID"

# アーキテクチャ判定
ARCH=$(uname -m)
echo "アーキテクチャ: $ARCH"

# LibreOfficeバージョン（最新安定版）
LIBREOFFICE_VERSION="24.8.4"
LIBREOFFICE_MINOR="24.8.4.2"

#------------------------------------------
# Amazon Linux 2023
#------------------------------------------
if [[ "$OS" == "amzn" && "$VERSION_ID" == "2023" ]]; then
    echo ""
    echo "=== Amazon Linux 2023 用のセットアップ ==="
    
    echo ""
    echo "[1/6] システム更新..."
    sudo dnf update -y

    echo ""
    echo "[2/6] Java 21 インストール..."
    sudo dnf install -y java-21-amazon-corretto-devel

    echo ""
    echo "[3/6] Maven インストール..."
    sudo dnf install -y maven

    echo ""
    echo "[4/6] LibreOffice 依存パッケージ インストール..."
    sudo dnf install -y \
        libXinerama \
        cups-libs \
        cairo \
        libX11 \
        wget \
        tar \
        gzip

    echo ""
    echo "[5/6] LibreOffice ダウンロード＆インストール..."
    
    # アーキテクチャに応じたダウンロードURL
    if [[ "$ARCH" == "x86_64" ]]; then
        LIBREOFFICE_URL="https://download.documentfoundation.org/libreoffice/stable/${LIBREOFFICE_VERSION}/rpm/x86_64/LibreOffice_${LIBREOFFICE_MINOR}_Linux_x86-64_rpm.tar.gz"
        LIBREOFFICE_DIR="LibreOffice_${LIBREOFFICE_MINOR}_Linux_x86-64_rpm"
    elif [[ "$ARCH" == "aarch64" ]]; then
        LIBREOFFICE_URL="https://download.documentfoundation.org/libreoffice/stable/${LIBREOFFICE_VERSION}/rpm/aarch64/LibreOffice_${LIBREOFFICE_MINOR}_Linux_aarch64_rpm.tar.gz"
        LIBREOFFICE_DIR="LibreOffice_${LIBREOFFICE_MINOR}_Linux_aarch64_rpm"
    else
        echo "未対応のアーキテクチャ: $ARCH"
        exit 1
    fi

    cd /tmp
    
    # ダウンロード（既存ファイルがあればスキップ）
    TARBALL="${LIBREOFFICE_DIR}.tar.gz"
    if [ ! -f "$TARBALL" ]; then
        echo "  ダウンロード中: $LIBREOFFICE_URL"
        wget -q --show-progress "$LIBREOFFICE_URL" -O "$TARBALL"
    else
        echo "  既存のファイルを使用: $TARBALL"
    fi
    
    # 展開
    echo "  展開中..."
    tar -xzf "$TARBALL"
    
    # RPMインストール
    echo "  RPMインストール中..."
    cd "$LIBREOFFICE_DIR/RPMS"
    sudo dnf install -y *.rpm
    
    # シンボリックリンク作成（libreoffice コマンドで呼び出せるように）
    LIBREOFFICE_BIN=$(ls /opt/libreoffice*/program/soffice 2>/dev/null | head -1)
    if [ -n "$LIBREOFFICE_BIN" ]; then
        LIBREOFFICE_DIR_PATH=$(dirname $(dirname "$LIBREOFFICE_BIN"))
        sudo ln -sf "$LIBREOFFICE_DIR_PATH/program/soffice" /usr/local/bin/libreoffice
        echo "  シンボリックリンク作成: /usr/local/bin/libreoffice -> $LIBREOFFICE_DIR_PATH/program/soffice"
    fi
    
    cd ~

    echo ""
    echo "[6/6] 日本語フォント & Python インストール..."
    sudo dnf install -y google-noto-sans-cjk-fonts || echo "  ※ google-noto-sans-cjk-fonts が見つかりません。代替をインストールします..."
    
    # 代替フォント
    sudo dnf install -y google-noto-fonts-common || true
    
    sudo dnf install -y python3 python3-pip
    pip3 install --user openpyxl python-docx pillow

#------------------------------------------
# Amazon Linux 2 (旧バージョン)
#------------------------------------------
elif [[ "$OS" == "amzn" && "$VERSION_ID" == "2" ]]; then
    echo ""
    echo "=== Amazon Linux 2 用のセットアップ ==="
    
    sudo yum update -y
    sudo amazon-linux-extras install -y java-openjdk11
    sudo yum install -y maven
    
    # Amazon Linux 2 は libreoffice トピックがある
    sudo amazon-linux-extras install -y libreoffice
    
    sudo yum install -y google-noto-sans-cjk-fonts
    sudo yum install -y python3 python3-pip
    pip3 install --user openpyxl python-docx pillow

#------------------------------------------
# Ubuntu / Debian系
#------------------------------------------
elif [[ "$OS" == "ubuntu" || "$OS" == "debian" ]]; then
    echo ""
    echo "=== Ubuntu/Debian 用のセットアップ ==="
    
    echo "[1/5] システム更新..."
    sudo apt update && sudo apt upgrade -y

    echo "[2/5] Java 21 インストール..."
    sudo apt install -y openjdk-21-jdk || sudo apt install -y openjdk-17-jdk

    echo "[3/5] Maven インストール..."
    sudo apt install -y maven

    echo "[4/5] LibreOffice インストール..."
    sudo apt install -y libreoffice-core libreoffice-calc libreoffice-writer

    echo "[5/5] 日本語フォント & Python インストール..."
    sudo apt install -y fonts-noto-cjk fonts-ipafont
    sudo apt install -y python3 python3-pip
    pip3 install --user openpyxl python-docx pillow

#------------------------------------------
# RHEL / CentOS / Fedora
#------------------------------------------
elif [[ "$OS" == "rhel" || "$OS" == "centos" || "$OS" == "fedora" ]]; then
    echo ""
    echo "=== RHEL/CentOS/Fedora 用のセットアップ ==="
    
    sudo dnf update -y
    sudo dnf install -y java-21-openjdk-devel || sudo dnf install -y java-17-openjdk-devel
    sudo dnf install -y maven
    sudo dnf install -y libreoffice-core libreoffice-calc libreoffice-writer
    sudo dnf install -y google-noto-sans-cjk-fonts
    sudo dnf install -y python3 python3-pip
    pip3 install --user openpyxl python-docx pillow

else
    echo "未対応のOS: $OS"
    echo "手動でJava 21, Maven, LibreOfficeをインストールしてください"
    exit 1
fi

echo ""
echo "=========================================="
echo "  インストール確認"
echo "=========================================="

echo ""
echo "Java:"
java -version 2>&1 | head -3

echo ""
echo "Maven:"
mvn -version 2>&1 | head -1

echo ""
echo "LibreOffice:"
if command -v libreoffice &> /dev/null; then
    libreoffice --version
elif [ -f /opt/libreoffice24.8/program/soffice ]; then
    /opt/libreoffice24.8/program/soffice --version
else
    echo "  LibreOfficeのパスを確認してください"
    ls /opt/ | grep -i libre || echo "  /opt/ にLibreOfficeが見つかりません"
fi

echo ""
echo "=========================================="
echo "  セットアップ完了！"
echo "=========================================="
echo ""
echo "次のステップ:"
echo "  1. cd pdf-converter-demo"
echo "  2. python3 create_test_files.py"
echo "  3. mvn clean package -DskipTests"
echo "  4. java -jar target/pdf-converter-demo-1.0.0.jar"
echo ""
