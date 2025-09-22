package com.watermark.service;

import com.watermark.config.WatermarkConfig;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.io.OutputStream;

public interface WatermarkService {

    void addWatermark(InputStream input, OutputStream output, String fileType, WatermarkConfig config) throws Exception;

    byte[] addWatermark(MultipartFile file, WatermarkConfig config) throws Exception;
}
