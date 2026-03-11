# PDF Converter with S3 Integration

Convert Excel, Word, and TIFF files to PDF and merge them into a single document.  
Files are downloaded from S3, processed, and the result is uploaded back to S3.

## Features

| Feature | Description |
|---------|-------------|
| Excel to PDF | Converts only a specific sheet (preserves all formatting) |
| Word to PDF | Converts DOCX/DOC files |
| TIFF to PDF | Rotates portrait images (height > width) 90 degrees clockwise |
| PDF Merge | Combines all converted PDFs into one document |
| S3 Integration | Downloads input files from S3, uploads result to S3 |

## Directory Structure

```
pdf-converter-s3/
├── pom.xml
├── README.md
└── src/main/java/com/example/
    ├── PdfConverterMain.java      # Main entry point
    ├── S3ClientUtil.java          # S3 client wrapper
    ├── LibreOfficeConverter.java  # Excel/Word to PDF conversion
    ├── ExcelSheetExtractor.java   # Extract specific sheet from Excel
    ├── TiffToPdfConverter.java    # TIFF to PDF (with rotation)
    └── PdfMerger.java             # Merge multiple PDFs
```

## Prerequisites

### 1. Java 21

```bash
java -version
# Should show version 21.x.x
```

### 2. Maven

```bash
mvn -version
```

### 3. LibreOffice (for Excel/Word conversion)

```bash
# Amazon Linux 2023
cd /tmp
wget https://ftp.osuosl.org/pub/tdf/libreoffice/stable/25.8.4/rpm/x86_64/LibreOffice_25.8.4_Linux_x86-64_rpm.tar.gz
tar -xzf LibreOffice_25.8.4_Linux_x86-64_rpm.tar.gz
cd LibreOffice_25.8.4*_rpm/RPMS
sudo dnf install -y *.rpm
sudo ln -sf /opt/libreoffice25.8/program/soffice /usr/local/bin/libreoffice

# Ubuntu
sudo apt install -y libreoffice-core libreoffice-calc libreoffice-writer
```

### 4. AWS Credentials

Configure AWS credentials using one of these methods:

```bash
# Method 1: Environment variables
export AWS_ACCESS_KEY_ID=your-access-key
export AWS_SECRET_ACCESS_KEY=your-secret-key

# Method 2: AWS CLI configuration
aws configure

# Method 3: IAM Role (recommended for EC2)
# Attach an IAM role with S3 access to your EC2 instance
```

## Build

```bash
cd pdf-converter-s3
mvn clean package -DskipTests
```

## Usage

### Command Line Arguments

```bash
java -jar target/pdf-converter-s3-1.0.0.jar \
  --bucket your-bucket-name \
  --region ap-northeast-1 \
  --input-prefix input/company001/ \
  --output-prefix output/company001/ \
  --sheet "Sheet1"
```

### Environment Variables

```bash
export S3_BUCKET_NAME=your-bucket-name
export AWS_REGION=ap-northeast-1
export S3_INPUT_PREFIX=input/company001/
export S3_OUTPUT_PREFIX=output/company001/
export TARGET_EXCEL_SHEET=Sheet1

java -jar target/pdf-converter-s3-1.0.0.jar
```

### Options

| Option | Environment Variable | Default | Description |
|--------|---------------------|---------|-------------|
| `--bucket` | `S3_BUCKET_NAME` | - | S3 bucket name (required) |
| `--region` | `AWS_REGION` | ap-northeast-1 | AWS region |
| `--input-prefix` | `S3_INPUT_PREFIX` | input/ | S3 prefix for input files |
| `--output-prefix` | `S3_OUTPUT_PREFIX` | output/ | S3 prefix for output files |
| `--sheet` | `TARGET_EXCEL_SHEET` | Sheet1 | Excel sheet name to convert |

## S3 Folder Structure Example

```
your-bucket/
├── input/
│   └── company001/
│       ├── data.xlsx         # Excel file (with multiple sheets)
│       ├── document.docx     # Word file
│       ├── image1.tiff       # TIFF file (landscape)
│       └── image2.tiff       # TIFF file (portrait - will be rotated)
└── output/
    └── company001/
        └── merged_document.pdf  # Output (created by this tool)
```

## Processing Flow

```
1. Download files from S3
   s3://bucket/input/company001/*.* -> local ./input/

2. Convert Excel to PDF
   - Extract specified sheet only
   - Convert using LibreOffice

3. Convert Word to PDF
   - Convert using LibreOffice

4. Convert TIFF to PDF
   - Rotate portrait images 90 degrees
   - Convert using PDFBox

5. Merge all PDFs
   - Combine into merged_document.pdf

6. Upload result to S3
   local ./output/merged_document.pdf -> s3://bucket/output/company001/
```

## Expected Output

```
========================================
  PDF Converter - Start
========================================
Configuration:
  Bucket: your-bucket-name
  Region: ap-northeast-1
  Input prefix: input/company001/
  Output prefix: output/company001/
  Target sheet: Sheet1

[1/5] Downloading files from S3...
  Found 4 file(s)
    Downloading: input/company001/data.xlsx
    Downloaded to: input/data.xlsx
    ...

[2/5] Converting Excel to PDF...
    Target sheet: Sheet1
    Sheet extracted (format preserved): temp/excel_single_sheet.xlsx
    Executing: libreoffice --headless --convert-to pdf ...
  OK: output/excel_output.pdf

[3/5] Converting Word to PDF...
    Executing: libreoffice --headless --convert-to pdf ...
  OK: output/word_output.pdf

[4/5] Converting TIFF to PDF...
    Processing: image1.tiff
    Page count: 1
    Page 1: Landscape (800x600) -> No rotation
  OK: output/image1.pdf
    Processing: image2.tiff
    Page count: 1
    Page 1: Portrait detected (600x800) -> Rotating 90 degrees
  OK: output/image2.pdf

[5/5] Merging PDFs...
    Merging 4 file(s)
    + excel_output.pdf
    + word_output.pdf
    + image1.pdf
    + image2.pdf
    Merge complete: output/merged_document.pdf
  OK: output/merged_document.pdf

[6/6] Uploading result to S3...
    Uploading: merged_document.pdf
    Uploaded to: s3://your-bucket-name/output/company001/merged_document.pdf

========================================
  PDF Converter - Complete
========================================
```

## IAM Policy (Minimum Required)

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "s3:GetObject",
        "s3:ListBucket"
      ],
      "Resource": [
        "arn:aws:s3:::your-bucket-name",
        "arn:aws:s3:::your-bucket-name/input/*"
      ]
    },
    {
      "Effect": "Allow",
      "Action": [
        "s3:PutObject"
      ],
      "Resource": [
        "arn:aws:s3:::your-bucket-name/output/*"
      ]
    }
  ]
}
```

## Troubleshooting

### LibreOffice not found

```bash
which libreoffice
# If not found, install LibreOffice (see Prerequisites)
```

### Font issues (Japanese characters)

```bash
# Install CJK fonts
sudo dnf install -y google-noto-sans-cjk-fonts  # Amazon Linux
sudo apt install -y fonts-noto-cjk              # Ubuntu
```

### S3 access denied

- Check AWS credentials are configured
- Verify IAM policy has required permissions
- Check bucket name and region are correct

### Java version mismatch

```bash
# Check Maven is using Java 21
export JAVA_HOME=/usr/lib/jvm/java-21-amazon-corretto
mvn -version
```
