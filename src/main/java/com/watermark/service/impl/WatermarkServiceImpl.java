package com.watermark.service.impl;

import com.watermark.config.WatermarkConfig;
import com.watermark.service.WatermarkService;
import com.watermark.service.handler.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Objects;

@Service
public class WatermarkServiceImpl implements WatermarkService {

    private final List<WatermarkHandler> handlers;

    @Autowired
    public WatermarkServiceImpl(List<WatermarkHandler> handlers) {
        this.handlers = handlers;
    }

    @Override
    public void addWatermark(InputStream input, OutputStream output, String fileName, WatermarkConfig config) throws Exception {
        WatermarkHandler handler = handlers.stream()
                .filter(h -> h.supports(fileName))
                .findFirst()
                .orElseThrow(() -> new UnsupportedOperationException("File type not supported: " + fileName));
        handler.addWatermark(input, output, config);
    }

    @Override
    public byte[] addWatermark(MultipartFile file, WatermarkConfig config) throws Exception {
        try (ByteArrayOutputStream output = new ByteArrayOutputStream();
             InputStream input = file.getInputStream()) {
            addWatermark(input, output, Objects.requireNonNull(file.getOriginalFilename()), config);
            return output.toByteArray();
        }
    }
}
