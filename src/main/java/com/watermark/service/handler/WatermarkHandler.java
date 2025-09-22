package com.watermark.service.handler;

import com.watermark.config.WatermarkConfig;

import java.io.InputStream;
import java.io.OutputStream;

public interface WatermarkHandler {

    void addWatermark(InputStream input, OutputStream output, WatermarkConfig config) throws Exception;

    boolean supports(String fileName);
}
