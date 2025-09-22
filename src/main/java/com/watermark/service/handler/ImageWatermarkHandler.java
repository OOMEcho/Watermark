package com.watermark.service.handler;

import com.watermark.config.WatermarkConfig;
import com.watermark.utils.FontUtils;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.io.OutputStream;

public class ImageWatermarkHandler {

    public void addWatermark(InputStream input, OutputStream output, WatermarkConfig config) throws Exception {
        BufferedImage sourceImage = ImageIO.read(input);
        if (sourceImage == null) {
            throw new IllegalArgumentException("无法读取输入图片");
        }

        BufferedImage watermarkedImage = createCompatibleImage(sourceImage);

        Graphics2D g2d = null;
        try {
            g2d = watermarkedImage.createGraphics();
            g2d.drawImage(sourceImage, 0, 0, null);
            setupRenderingHints(g2d);
            addTextWatermark(g2d, config, sourceImage.getWidth(), sourceImage.getHeight());
        } finally {
            if (g2d != null) {
                g2d.dispose();
            }
        }

        // 始终输出 PNG 保留透明度
        ImageIO.write(watermarkedImage, "png", output);
    }

    /**
     * 创建与源图兼容的 BufferedImage
     */
    private BufferedImage createCompatibleImage(BufferedImage source) {
        int type = source.getTransparency() == Transparency.OPAQUE ?
                BufferedImage.TYPE_INT_RGB : BufferedImage.TYPE_INT_ARGB;
        return new BufferedImage(source.getWidth(), source.getHeight(), type);
    }

    /**
     * 开启抗锯齿
     */
    private void setupRenderingHints(Graphics2D g2d) {
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
    }

    /**
     * 绘制文字水印
     */
    private void addTextWatermark(Graphics2D g2d, WatermarkConfig config, int width, int height) {
        g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, config.getOpacity()));
        g2d.setColor(config.getColor());
        g2d.setFont(FontUtils.getSimSunFont(config.getFontSize()));

        FontMetrics fm = g2d.getFontMetrics();
        int textWidth = fm.stringWidth(config.getText());
        int textHeight = fm.getHeight();

        // 动态边距和步长
        int marginX = Math.max(20, width / 50);
        int marginY = Math.max(20, height / 50);
        int xStep = textWidth + marginX;
        int yStep = textHeight + marginY;

        if (config.getPosition() == WatermarkConfig.Position.DIAGONAL) {
            // 斜线平铺
            int cols = (int) Math.ceil((double) width / xStep);
            int rows = (int) Math.ceil((double) height / yStep);
            for (int row = 0; row <= rows; row++) {
                for (int col = 0; col <= cols; col++) {
                    int x = col * xStep;
                    int y = row * yStep + textHeight;

                    drawRotatedString(g2d, config, x, y, textWidth, textHeight);
                }
            }
        } else {
            // 单条水印
            Point position = calculatePosition(config.getPosition(), width, height, textWidth, textHeight, marginX, marginY);
            drawRotatedString(g2d, config, position.x, position.y, textWidth, textHeight);
        }
    }

    /**
     * 绘制单条文字并旋转
     */
    private void drawRotatedString(Graphics2D g2d, WatermarkConfig config, int x, int y, int textWidth, int textHeight) {
        AffineTransform original = g2d.getTransform();
        if (config.getRotation() != 0) {
            g2d.rotate(Math.toRadians(config.getRotation()), x + textWidth / 2.0, y - textHeight / 2.0);
        }
        g2d.drawString(config.getText(), x, y);
        g2d.setTransform(original);
    }

    /**
     * 计算单条水印位置
     */
    private Point calculatePosition(WatermarkConfig.Position position, int width, int height,
                                    int textWidth, int textHeight, int marginX, int marginY) {
        switch (position) {
            case TOP_LEFT:
                return new Point(marginX, marginY + textHeight);
            case TOP_RIGHT:
                return new Point(width - textWidth - marginX, marginY + textHeight);
            case BOTTOM_LEFT:
                return new Point(marginX, height - marginY);
            case BOTTOM_RIGHT:
                return new Point(width - textWidth - marginX, height - marginY);
            case CENTER:
            default:
                return new Point((width - textWidth) / 2, (height + textHeight) / 2);
        }
    }
}
