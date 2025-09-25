package com.watermark.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.awt.*;

@Data
@Configuration
@ConfigurationProperties(prefix = "watermark")
public class WatermarkConfig {

    private String text = "aegis";

    private Float opacity = 0.3f;

    private Integer fontSize = 40;

    private String color = "GRAY";

    private Float rotation = 45f;

    private Position position = Position.DIAGONAL;

    public Color getColorObject() {
        try {
            return (Color) Color.class.getField(color.toUpperCase()).get(null);
        } catch (Exception e) {
            return Color.GRAY;
        }
    }

    public enum Position {
        CENTER, TOP_LEFT, TOP_RIGHT, BOTTOM_LEFT, BOTTOM_RIGHT, DIAGONAL
    }

}
