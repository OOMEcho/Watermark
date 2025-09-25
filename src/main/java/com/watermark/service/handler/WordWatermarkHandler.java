package com.watermark.service.handler;

import com.watermark.config.WatermarkConfig;
import com.watermark.utils.FontUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Component;

import java.awt.Font;
import java.io.InputStream;
import java.io.OutputStream;

@Slf4j
@Component
public class WordWatermarkHandler implements WatermarkHandler {

    @Override
    public void addWatermark(InputStream input, OutputStream output, WatermarkConfig config) throws Exception {
        XWPFDocument doc = new XWPFDocument(input);

        addBackgroundWatermark(doc, config);

        doc.write(output);
        doc.close();
    }

    @Override
    public boolean supports(String fileName) {
        return fileName.toLowerCase().endsWith(".docx") || fileName.toLowerCase().endsWith(".doc");
    }

    private void addBackgroundWatermark(XWPFDocument doc, WatermarkConfig config) {
        XWPFHeaderFooterPolicy policy = doc.getHeaderFooterPolicy();
        if (policy == null) {
            policy = doc.createHeaderFooterPolicy();
        }

        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        if (config.getPosition() == WatermarkConfig.Position.DIAGONAL) {
            createDiagonalWatermark(header, config);
        } else {
            createSingleWatermark(header, config);
        }
    }

    /**
     * 创建单个位置的背景水印
     */
    private void createSingleWatermark(XWPFHeader header, WatermarkConfig config) {
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = paragraph.createRun();

        // 使用FontUtils获取支持中文的字体
        Font chineseFont = FontUtils.getChineseFont(config.getFontSize());
        String fontFamily = chineseFont.getFontName();
        log.debug("使用字体: {} 来渲染水印文本", fontFamily);

        run.setFontFamily(fontFamily);
        run.setFontSize(config.getFontSize());
        run.setText(config.getText());

        String colorHex = String.format("%06X", config.getColorObject().getRGB() & 0xFFFFFF);
        run.setColor(colorHex);

        try {
            String watermarkXml = createBackgroundWatermarkXml(config, fontFamily, colorHex);
            run.getCTR().set(org.apache.xmlbeans.XmlObject.Factory.parse(watermarkXml));
            log.info("成功创建背景水印，位置: {}", config.getPosition());
        } catch (Exception e) {
            log.error("VML背景水印创建失败: {}", e.getMessage());
        }
    }

    /**
     * 创建对角线平铺背景水印
     */
    private void createDiagonalWatermark(XWPFHeader header, WatermarkConfig config) {
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);

        Font chineseFont = FontUtils.getChineseFont(config.getFontSize());
        String fontFamily = chineseFont.getFontName();
        String colorHex = String.format("%06X", config.getColorObject().getRGB() & 0xFFFFFF);

        // 创建多个水印实现平铺效果
        int[] positions = {
            // 覆盖整页的网格位置 (相对位置百分比)
            15, 25,   35, 25,   55, 25,   75, 25,   // 第一行
            5, 45,    25, 45,   45, 45,   65, 45,   85, 45, // 第二行
            15, 65,   35, 65,   55, 65,   75, 65,   // 第三行
            5, 85,    25, 85,   45, 85,   65, 85,   85, 85  // 第四行
        };

        for (int i = 0; i < positions.length; i += 2) {
            XWPFRun run = paragraph.createRun();
            run.setFontFamily(fontFamily);
            run.setFontSize(config.getFontSize());
            run.setText(config.getText());
            run.setColor(colorHex);

            try {
                String watermarkXml = createDiagonalWatermarkXml(config, fontFamily, colorHex,
                                                                positions[i], positions[i + 1]);
                run.getCTR().set(org.apache.xmlbeans.XmlObject.Factory.parse(watermarkXml));
            } catch (Exception e) {
                log.warn("创建对角线水印位置 ({}, {}) 失败: {}", positions[i], positions[i + 1], e.getMessage());
            }
        }

        log.info("成功创建对角线平铺背景水印，共 {} 个水印", positions.length / 2);
    }

    /**
     * 创建单个背景水印的XML
     */
    private String createBackgroundWatermarkXml(WatermarkConfig config, String fontFamily, String colorHex) {
        String rotation = String.valueOf((int) config.getRotation().floatValue());
        String[] position = getWatermarkPosition(config.getPosition());
        String horizontalPos = position[0];
        String verticalPos = position[1];

        // 计算透明度（0-1转换为0-65535）
        int alpha = (int) (config.getOpacity() * 65535);

        return "<w:pict xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
               "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" id=\"watermark\" type=\"#_x0000_t136\" " +
               "style=\"position:absolute;width:500pt;height:150pt;z-index:-251658240;" +
               "mso-position-horizontal:" + horizontalPos + ";" +
               "mso-position-vertical:" + verticalPos + ";" +
               "mso-position-horizontal-relative:margin;" +
               "mso-position-vertical-relative:margin;" +
               "rotation:" + rotation + "\" " +
               "fillcolor=\"" + colorHex + "\" stroked=\"false\" " +
               "opacity=\"" + alpha + "\">" +
               "<v:textpath on=\"true\" fitpath=\"true\" " +
               "string=\"" + escapeXmlText(config.getText()) + "\" " +
               "style=\"font-family:" + escapeXmlText(fontFamily) + ";" +
               "font-size:" + config.getFontSize() + "pt;" +
               "color:" + colorHex + ";" +
               "font-weight:bold\"/>" +
               "</v:shape>" +
               "</w:pict>";
    }

    /**
     * 创建对角线平铺水印的XML
     */
    private String createDiagonalWatermarkXml(WatermarkConfig config, String fontFamily,
                                            String colorHex, int leftPercent, int topPercent) {
        String rotation = String.valueOf((int) config.getRotation().floatValue());

        // 计算透明度
        int alpha = (int) (config.getOpacity() * 65535);

        return "<w:pict xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
               "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" type=\"#_x0000_t136\" " +
               "style=\"position:absolute;width:200pt;height:50pt;z-index:-251658240;" +
               "left:" + leftPercent + "%;" +
               "top:" + topPercent + "%;" +
               "rotation:" + rotation + "\" " +
               "fillcolor=\"" + colorHex + "\" stroked=\"false\" " +
               "opacity=\"" + alpha + "\">" +
               "<v:textpath on=\"true\" fitpath=\"true\" " +
               "string=\"" + escapeXmlText(config.getText()) + "\" " +
               "style=\"font-family:" + escapeXmlText(fontFamily) + ";" +
               "font-size:" + (config.getFontSize() * 0.8) + "pt;" +
               "color:" + colorHex + "\"/>" +
               "</v:shape>" +
               "</w:pict>";
    }

    /**
     * 根据位置枚举返回VML位置参数
     */
    private String[] getWatermarkPosition(WatermarkConfig.Position position) {
        switch (position) {
            case TOP_LEFT:
                return new String[]{"left", "top"};
            case TOP_RIGHT:
                return new String[]{"right", "top"};
            case BOTTOM_LEFT:
                return new String[]{"left", "bottom"};
            case BOTTOM_RIGHT:
                return new String[]{"right", "bottom"};
            case CENTER:
            default:
                return new String[]{"center", "center"};
        }
    }

    /**
     * 转义XML文本中的特殊字符
     */
    private String escapeXmlText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }
}
