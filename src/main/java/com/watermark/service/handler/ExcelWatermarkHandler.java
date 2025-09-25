package com.watermark.service.handler;

import com.watermark.config.WatermarkConfig;
import com.watermark.utils.FontUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

@Slf4j
@Component
public class ExcelWatermarkHandler implements WatermarkHandler {

    /**
     * 给 Excel 添加水印
     */
    @Override
    public void addWatermark(InputStream input, OutputStream output, WatermarkConfig config) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(input)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                addTextWatermarkToSheet(sheet, config, workbook);
            }
            workbook.write(output);
        }
    }

    @Override
    public boolean supports(String fileName) {
        return fileName.toLowerCase().endsWith(".xlsx") || fileName.toLowerCase().endsWith(".xls");
    }

    /**
     * 在单个 Sheet 添加文字水印
     */
    private void addTextWatermarkToSheet(XSSFSheet sheet, WatermarkConfig config, XSSFWorkbook workbook) {
        try {
            Dimension sheetSize = calculateSheetPixelSize(sheet);
            BufferedImage watermarkImage = createTextImage(config, sheetSize);

            try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                ImageIO.write(watermarkImage, "png", baos);

                int pictureIdx = workbook.addPicture(baos.toByteArray(), Workbook.PICTURE_TYPE_PNG);
                XSSFDrawing drawing = sheet.createDrawingPatriarch();

                XSSFClientAnchor anchor = createDynamicAnchor(sheet, config);
                XSSFPicture picture = drawing.createPicture(anchor, pictureIdx);
                picture.getCTPicture().getNvPicPr().getCNvPicPr().addNewPicLocks().setNoChangeAspect(true);
            }
        } catch (IOException e) {
            log.warn("水印图片生成失败，跳过此 Sheet，原因: {}", e.getMessage(), e);
        }
    }

    /**
     * 创建水印图片
     */
    private BufferedImage createTextImage(WatermarkConfig config, Dimension sheetSize) {
        int width = sheetSize.width;
        int height = sheetSize.height;

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = image.createGraphics();

        try {
            // 抗锯齿
            g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2d.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

            // 清空背景
            g2d.setComposite(AlphaComposite.Clear);
            g2d.fillRect(0, 0, width, height);

            // 透明度和颜色
            g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, config.getOpacity()));
            g2d.setColor(config.getColorObject());
            g2d.setFont(FontUtils.getChineseFont(config.getFontSize()));

            // 旋转
            if (config.getRotation() != 0) {
                g2d.rotate(Math.toRadians(config.getRotation()), (double) width / 2, (double) height / 2);
            }

            // 绘制文字平铺
            FontMetrics fm = g2d.getFontMetrics();
            int textWidth = fm.stringWidth(config.getText());
            int textHeight = fm.getHeight();

            for (int y = 0; y < height; y += textHeight * 3) {
                for (int x = 0; x < width; x += textWidth * 2) {
                    g2d.drawString(config.getText(), x, y);
                }
            }
        } finally {
            g2d.dispose();
        }

        return image;
    }

    /**
     * 动态计算 Sheet 在像素上的宽高
     */
    private Dimension calculateSheetPixelSize(XSSFSheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return new Dimension(800, 600); // 空表默认大小
        }

        // 计算最大列
        int maxCol = getLastColumnNum(sheet);

        // 总宽度
        int totalWidth = 0;
        for (int col = 0; col < maxCol; col++) {
            totalWidth += (int) sheet.getColumnWidthInPixels(col);
        }

        // 总高度
        int totalHeight = 0;
        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                totalHeight += (int) (row.getHeightInPoints() * 96 / 72);
            }
        }

        // 增加额外 padding，防止最后一行漏出
        int padding = 50;
        return new Dimension(Math.max(totalWidth, 800), Math.max(totalHeight + padding, 600));
    }

    /**
     * 计算 Sheet 的最大列数
     */
    private int getLastColumnNum(XSSFSheet sheet) {
        int maxCol = 0;
        for (Row row : sheet) {
            if (row.getLastCellNum() > maxCol) {
                maxCol = row.getLastCellNum();
            }
        }
        return maxCol;
    }

    /**
     * 创建水印位置 Anchor
     */
    private XSSFClientAnchor createDynamicAnchor(XSSFSheet sheet, WatermarkConfig config) {
        int lastRowNum = sheet.getLastRowNum();
        int lastColNum = getLastColumnNum(sheet);

        XSSFClientAnchor anchor = new XSSFClientAnchor();

        switch (config.getPosition()) {
            case TOP_LEFT:
                anchor.setCol1(0);
                anchor.setRow1(0);
                anchor.setCol2(Math.min(5, lastColNum));
                anchor.setRow2(Math.min(8, lastRowNum));
                break;
            case TOP_RIGHT:
                anchor.setCol1(Math.max(0, lastColNum - 5));
                anchor.setRow1(0);
                anchor.setCol2(lastColNum);
                anchor.setRow2(Math.min(8, lastRowNum));
                break;
            case BOTTOM_LEFT:
                anchor.setCol1(0);
                anchor.setRow1(Math.max(0, lastRowNum - 8));
                anchor.setCol2(Math.min(5, lastColNum));
                anchor.setRow2(lastRowNum);
                break;
            case BOTTOM_RIGHT:
                anchor.setCol1(Math.max(0, lastColNum - 5));
                anchor.setRow1(Math.max(0, lastRowNum - 8));
                anchor.setCol2(lastColNum);
                anchor.setRow2(lastRowNum);
                break;
            case CENTER:
                int centerCol = lastColNum / 2;
                int centerRow = lastRowNum / 2;
                anchor.setCol1(Math.max(0, centerCol - 3));
                anchor.setRow1(Math.max(0, centerRow - 4));
                anchor.setCol2(Math.min(lastColNum, centerCol + 3));
                anchor.setRow2(Math.min(lastRowNum, centerRow + 4));
                break;
            default:
                anchor.setCol1(0);
                anchor.setRow1(0);
                anchor.setCol2(lastColNum);
                anchor.setRow2(lastRowNum + 1);
        }

        return anchor;
    }
}
