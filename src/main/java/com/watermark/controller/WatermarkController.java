package com.watermark.controller;

import com.watermark.config.WatermarkConfig;
import com.watermark.service.WatermarkService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.awt.*;

@RestController
@RequestMapping("/api/watermark")
public class WatermarkController {

    @Autowired
    private WatermarkService watermarkService;

    @PostMapping("/add")
    public ResponseEntity<byte[]> addWatermark(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "text", defaultValue = "aegis") String watermarkText,
            @RequestParam(value = "opacity", defaultValue = "0.3") Float opacity,
            @RequestParam(value = "fontSize", defaultValue = "40") Integer fontSize,
            @RequestParam(value = "rotation", defaultValue = "45") Float rotation,
            @RequestParam(value = "position", defaultValue = "DIAGONAL") String position) {

        try {
            WatermarkConfig config = new WatermarkConfig();
            config.setText(watermarkText);
            config.setOpacity(opacity);
            config.setFontSize(fontSize);
            config.setRotation(rotation);
            config.setPosition(WatermarkConfig.Position.valueOf(position.toUpperCase()));
            config.setColor(Color.GRAY);

            byte[] result = watermarkService.addWatermark(file, config);

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
