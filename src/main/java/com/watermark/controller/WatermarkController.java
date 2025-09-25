package com.watermark.controller;

import com.watermark.config.WatermarkConfig;
import com.watermark.service.WatermarkService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/watermark")
public class WatermarkController {

    @Autowired
    private WatermarkService watermarkService;

    @Autowired
    private WatermarkConfig watermarkConfig;

    @PostMapping("/add")
    public ResponseEntity<byte[]> addWatermark(@RequestParam("file") MultipartFile file) {

        try {
            byte[] result = watermarkService.addWatermark(file, watermarkConfig);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment",
                "watermarked_" + file.getOriginalFilename());

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(result);

        } catch (Exception e) {
            return ResponseEntity.badRequest().build();
        }
    }
}
