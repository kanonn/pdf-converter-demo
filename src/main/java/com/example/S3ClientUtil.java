package com.example;

import software.amazon.awssdk.core.sync.RequestBody;
import software.amazon.awssdk.regions.Region;
import software.amazon.awssdk.services.s3.S3Client;
import software.amazon.awssdk.services.s3.model.*;

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

    public S3ClientUtil(String bucketName, String region) {
        this.bucketName = bucketName;
        this.s3Client = S3Client.builder()
                .region(Region.of(region))
                .build();
    }

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

    public void downloadFile(String s3Key, Path localPath) throws Exception {
        System.out.println("    Downloading: " + s3Key);
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

    public void close() {
        if (s3Client != null) {
            s3Client.close();
        }
    }
}
