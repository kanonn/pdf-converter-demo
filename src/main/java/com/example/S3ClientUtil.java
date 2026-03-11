package com.example;

import software.amazon.awssdk.core.sync.RequestBody;
import software.amazon.awssdk.regions.Region;
import software.amazon.awssdk.services.s3.S3Client;
import software.amazon.awssdk.services.s3.model.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

/**
 * S3 client utility for file upload/download operations.
 */
public class S3ClientUtil {

    private final S3Client s3Client;
    private final String bucketName;

    /**
     * Constructor with bucket name and region.
     *
     * @param bucketName S3 bucket name
     * @param region     AWS region (e.g., "ap-northeast-1")
     */
    public S3ClientUtil(String bucketName, String region) {
        this.bucketName = bucketName;
        this.s3Client = S3Client.builder()
                .region(Region.of(region))
                .build();
    }

    /**
     * Constructor with existing S3Client.
     *
     * @param s3Client   S3Client instance
     * @param bucketName S3 bucket name
     */
    public S3ClientUtil(S3Client s3Client, String bucketName) {
        this.s3Client = s3Client;
        this.bucketName = bucketName;
    }

    /**
     * List all objects under a specific prefix (folder).
     *
     * @param prefix S3 key prefix (e.g., "input/company001/")
     * @return List of S3 object keys
     */
    public List<String> listObjects(String prefix) {
        List<String> keys = new ArrayList<>();

        ListObjectsV2Request request = ListObjectsV2Request.builder()
                .bucket(bucketName)
                .prefix(prefix)
                .build();

        ListObjectsV2Response response;
        do {
            response = s3Client.listObjectsV2(request);
            for (S3Object object : response.contents()) {
                // Skip folder markers (keys ending with /)
                if (!object.key().endsWith("/")) {
                    keys.add(object.key());
                }
            }
            request = request.toBuilder()
                    .continuationToken(response.nextContinuationToken())
                    .build();
        } while (response.isTruncated());

        return keys;
    }

    /**
     * Download a file from S3 to local path.
     *
     * @param s3Key     S3 object key
     * @param localPath Local file path to save
     * @throws IOException if download fails
     */
    public void downloadFile(String s3Key, Path localPath) throws IOException {
        System.out.println("    Downloading: " + s3Key);

        // Create parent directories if not exist
        Files.createDirectories(localPath.getParent());

        GetObjectRequest request = GetObjectRequest.builder()
                .bucket(bucketName)
                .key(s3Key)
                .build();

        try (InputStream inputStream = s3Client.getObject(request)) {
            Files.copy(inputStream, localPath, StandardCopyOption.REPLACE_EXISTING);
        }

        System.out.println("    Downloaded to: " + localPath);
    }

    /**
     * Upload a file from local path to S3.
     *
     * @param localPath Local file path
     * @param s3Key     S3 object key
     */
    public void uploadFile(Path localPath, String s3Key) {
        System.out.println("    Uploading: " + localPath.getFileName());

        PutObjectRequest request = PutObjectRequest.builder()
                .bucket(bucketName)
                .key(s3Key)
                .build();

        s3Client.putObject(request, RequestBody.fromFile(localPath.toFile()));

        System.out.println("    Uploaded to: s3://" + bucketName + "/" + s3Key);
    }

    /**
     * Upload a file with content type specification.
     *
     * @param localPath   Local file path
     * @param s3Key       S3 object key
     * @param contentType MIME content type (e.g., "application/pdf")
     */
    public void uploadFile(Path localPath, String s3Key, String contentType) {
        System.out.println("    Uploading: " + localPath.getFileName());

        PutObjectRequest request = PutObjectRequest.builder()
                .bucket(bucketName)
                .key(s3Key)
                .contentType(contentType)
                .build();

        s3Client.putObject(request, RequestBody.fromFile(localPath.toFile()));

        System.out.println("    Uploaded to: s3://" + bucketName + "/" + s3Key);
    }

    /**
     * Check if an object exists in S3.
     *
     * @param s3Key S3 object key
     * @return true if exists
     */
    public boolean exists(String s3Key) {
        try {
            HeadObjectRequest request = HeadObjectRequest.builder()
                    .bucket(bucketName)
                    .key(s3Key)
                    .build();
            s3Client.headObject(request);
            return true;
        } catch (NoSuchKeyException e) {
            return false;
        }
    }

    /**
     * Delete an object from S3.
     *
     * @param s3Key S3 object key
     */
    public void deleteObject(String s3Key) {
        DeleteObjectRequest request = DeleteObjectRequest.builder()
                .bucket(bucketName)
                .key(s3Key)
                .build();
        s3Client.deleteObject(request);
    }

    /**
     * Get the bucket name.
     *
     * @return bucket name
     */
    public String getBucketName() {
        return bucketName;
    }

    /**
     * Close the S3 client.
     */
    public void close() {
        if (s3Client != null) {
            s3Client.close();
        }
    }
}
