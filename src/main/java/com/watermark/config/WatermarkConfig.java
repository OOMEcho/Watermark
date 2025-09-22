package com.watermark.config;

import lombok.Data;

import java.awt.*;

@Data
public class WatermarkConfig {

    private String text;

    private Float opacity = 0.3f;

    private Integer fontSize = 40;

    private Color color = Color.GRAY;

    private Float rotation = 45f;

    private Position position = Position.DIAGONAL;

    public enum Position {
        CENTER, TOP_LEFT, TOP_RIGHT, BOTTOM_LEFT, BOTTOM_RIGHT, DIAGONAL
    }

}
