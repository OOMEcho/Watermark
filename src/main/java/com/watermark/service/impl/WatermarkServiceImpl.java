package com.watermark.service.impl;

import com.watermark.config.WatermarkConfig;
import com.watermark.service.WatermarkService;
import com.watermark.service.handler.ExcelWatermarkHandler;
import com.watermark.service.handler.ImageWatermarkHandler;
import com.watermark.service.handler.PdfWatermarkHandler;
import com.watermark.service.handler.WordWatermarkHandler;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;

@Service
public class WatermarkServiceImpl implements WatermarkService {

    private final PdfWatermarkHandler pdfHandler;
    private final ImageWatermarkHandler imageHandler;
    private final WordWatermarkHandler wordHandler;
    private final ExcelWatermarkHandler excelHandler;

    public WatermarkServiceImpl() {
        this.pdfHandler = new PdfWatermarkHandler();
        this.imageHandler = new ImageWatermarkHandler();
        this.wordHandler = new WordWatermarkHandler();
        this.excelHandler = new ExcelWatermarkHandler();
    }

    @Override
    public void addWatermark(InputStream input, OutputStream output, String fileType, WatermarkConfig config) throws Exception {
        String lowerType = fileType.toLowerCase();

        if (lowerType.endsWith(".pdf")) {
            pdfHandler.addWatermark(input, output, config);
        } else if (isImageFile(lowerType)) {
            imageHandler.addWatermark(input, output, config);
        } else if (lowerType.endsWith(".docx") || lowerType.endsWith(".doc")) {
            wordHandler.addWatermark(input, output, config);
        } else if (lowerType.endsWith(".xlsx") || lowerType.endsWith(".xls")) {
            excelHandler.addWatermark(input, output, config);
        } else {
            throw new UnsupportedOperationException("File type not supported: " + fileType);
        }
    }

    @Override
    public byte[] addWatermark(MultipartFile file, WatermarkConfig config) throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        addWatermark(file.getInputStream(), output, Objects.requireNonNull(file.getOriginalFilename()), config);
        return output.toByteArray();
    }

    private boolean isImageFile(String fileName) {
        return fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") ||
               fileName.endsWith(".png") || fileName.endsWith(".gif") ||
               fileName.endsWith(".bmp");
    }
}
