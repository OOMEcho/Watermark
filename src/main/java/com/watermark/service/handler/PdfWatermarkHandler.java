package com.watermark.service.handler;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.*;
import com.watermark.config.WatermarkConfig;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

@Slf4j
@Component
public class PdfWatermarkHandler implements WatermarkHandler {

    @Override
    public void addWatermark(InputStream input, OutputStream output, WatermarkConfig config) throws Exception {
        PdfReader reader = null;
        PdfStamper stamper = null;

        try {
            reader = new PdfReader(input);
            stamper = new PdfStamper(reader, output);

            int pageCount = reader.getNumberOfPages();

            BaseFont baseFont = loadFont();

            for (int i = 1; i <= pageCount; i++) {
                Rectangle pageSize = reader.getPageSizeWithRotation(i);
                PdfContentByte content = stamper.getOverContent(i);
                addWatermarkToPage(content, baseFont, pageSize, config);
            }

        } finally {
            if (stamper != null) {
                try {
                    stamper.close();
                } catch (Exception e) { /* 忽略关闭异常 */ }
            }
            if (reader != null) {
                try {
                    reader.close();
                } catch (Exception e) { /* 忽略关闭异常 */ }
            }
        }
    }

    @Override
    public boolean supports(String fileName) {
        return fileName.toLowerCase().endsWith(".pdf");
    }

    private BaseFont loadFont() {
        try {
            InputStream fontStream = PdfWatermarkHandler.class.getResourceAsStream("/fonts/simsun.ttf");
            if (fontStream == null) {
                throw new IOException("字体文件未找到: " + "/fonts/simsun.ttf");
            }
            byte[] fontBytes = toByteArray(fontStream);
            return BaseFont.createFont("simsun.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, true, fontBytes, null);
        } catch (Exception e) {
            // 回退到内置 Helvetica 字体
            try {
                return BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
            } catch (Exception ex) {
                throw new RuntimeException("加载字体失败", ex);
            }
        }
    }

    private void addWatermarkToPage(PdfContentByte content, BaseFont baseFont, Rectangle pageSize, WatermarkConfig config) {
        content.saveState();
        PdfGState gs = new PdfGState();
        gs.setFillOpacity(config.getOpacity());
        gs.setStrokeOpacity(config.getOpacity());
        content.setGState(gs);
        content.setColorFill(new BaseColor(config.getColorObject().getRGB()));
        content.beginText();
        content.setFontAndSize(baseFont, config.getFontSize());

        if (config.getPosition() == WatermarkConfig.Position.DIAGONAL) {
            addDiagonalWatermark(content, baseFont, pageSize, config);
        } else {
            float[] pos = calculatePosition(config.getPosition(), pageSize);
            content.showTextAligned(Element.ALIGN_CENTER, config.getText(), pos[0], pos[1], config.getRotation());
        }

        content.endText();
        content.restoreState();
    }

    /**
     * 计算单条水印位置
     */
    private float[] calculatePosition(WatermarkConfig.Position position, Rectangle pageSize) {
        float x, y;
        switch (position) {
            case TOP_LEFT:
                x = 100;
                y = pageSize.getHeight() - 100;
                break;
            case TOP_RIGHT:
                x = pageSize.getWidth() - 100;
                y = pageSize.getHeight() - 100;
                break;
            case BOTTOM_LEFT:
                x = 100;
                y = 100;
                break;
            case BOTTOM_RIGHT:
                x = pageSize.getWidth() - 100;
                y = 100;
                break;
            case CENTER:
            default:
                x = pageSize.getWidth() / 2;
                y = pageSize.getHeight() / 2;
                break;
        }
        return new float[]{x, y};
    }

    /**
     * 斜线平铺水印，动态计算步长，覆盖整页
     */
    private void addDiagonalWatermark(PdfContentByte content, BaseFont baseFont, Rectangle pageSize, WatermarkConfig config) {
        float textWidth = baseFont.getWidthPoint(config.getText(), config.getFontSize());
        float textHeight = config.getFontSize();
        float marginX = 50;
        float marginY = 50;
        float stepX = textWidth + marginX;
        float stepY = textHeight + marginY;

        for (float y = textHeight; y < pageSize.getHeight(); y += stepY) {
            for (float x = 0; x < pageSize.getWidth(); x += stepX) {
                content.showTextAligned(Element.ALIGN_CENTER, config.getText(), x, y, config.getRotation());
            }
        }
    }

    private byte[] toByteArray(InputStream in) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[8192];
        int len;
        while ((len = in.read(buffer)) != -1) {
            baos.write(buffer, 0, len);
        }
        return baos.toByteArray();
    }
}
