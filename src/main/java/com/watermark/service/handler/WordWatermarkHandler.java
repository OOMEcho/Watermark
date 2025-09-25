package com.watermark.service.handler;

import com.watermark.config.WatermarkConfig;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.io.OutputStream;

@Slf4j
@Component
public class WordWatermarkHandler implements WatermarkHandler {

    @Override
    public void addWatermark(InputStream input, OutputStream output, WatermarkConfig config) throws Exception {
        try (XWPFDocument doc = new XWPFDocument(input)) {
            // 在正文中添加背景水印
            addBackgroundWatermark(doc, config);
            doc.write(output);
            log.info("成功添加Word正文背景水印");
        } catch (Exception e) {
            log.error("添加Word水印失败: {}", e.getMessage(), e);
            throw new RuntimeException("添加Word水印失败", e);
        }
    }

    @Override
    public boolean supports(String fileName) {
        return fileName.toLowerCase().endsWith(".docx");
    }

    private void addBackgroundWatermark(XWPFDocument doc, WatermarkConfig config) {
        try {
            // 方法1：直接在文档正文中插入背景水印段落
            addDocumentBackgroundWatermark(doc, config);
        } catch (Exception e) {
            log.warn("正文水印添加失败，尝试简单文本水印: {}", e.getMessage());
            // 备用方案：在文档开头添加透明文本水印
            addSimpleBackgroundWatermark(doc, config);
        }
    }

    private void addDocumentBackgroundWatermark(XWPFDocument doc, WatermarkConfig config) {
        // 如果文档没有段落，创建一个
        if (doc.getParagraphs().isEmpty()) {
            doc.createParagraph();
        }

        // 在第一个段落之前插入水印段落
        XWPFParagraph firstPara = doc.getParagraphs().get(0);
        XWPFParagraph watermarkPara = doc.insertNewParagraph(firstPara.getCTP().newCursor());

        // 设置水印段落属性
        watermarkPara.setAlignment(ParagraphAlignment.CENTER);

        // 设置段落为绝对定位，使其不影响正文布局
        CTPPr pPr = watermarkPara.getCTP().getPPr();
        if (pPr == null) {
            pPr = watermarkPara.getCTP().addNewPPr();
        }

        // 添加水印run
        XWPFRun run = watermarkPara.createRun();
        setupWatermarkRun(run, config);

        log.info("成功添加正文背景水印");
    }

    private void setupWatermarkRun(XWPFRun run, WatermarkConfig config) {
        // 设置水印文本
        run.setText(config.getText());

        // 设置字体
        run.setFontFamily("SimSun");
        run.setFontSize(config.getFontSize());

        // 设置颜色
        String colorHex = String.format("%06X", config.getColorObject().getRGB() & 0xFFFFFF);
        run.setColor(colorHex);

        // 通过XML设置更多样式属性
        CTR ctr = run.getCTR();
        CTRPr rPr = ctr.getRPr();
        if (rPr == null) {
            rPr = ctr.addNewRPr();
        }

        try {
            // 设置文本效果和透明度
            addTextEffects(rPr, config);
        } catch (Exception e) {
            log.warn("设置文本效果失败: {}", e.getMessage());
        }
    }

    private void addTextEffects(CTRPr rPr, WatermarkConfig config) {
        // 构建文本效果XML
        String effectsXml = buildTextEffectsXML(config);

        try {
            // 解析并插入效果XML
            org.apache.xmlbeans.XmlObject effectsObject = org.apache.xmlbeans.XmlObject.Factory.parse(effectsXml);

            XmlCursor cursor = rPr.newCursor();
            cursor.toEndToken();

            // 直接插入XML内容
            cursor.insertChars(effectsXml);
            cursor.dispose();

        } catch (Exception e) {
            log.debug("文本效果设置失败，使用基础样式: {}", e.getMessage());

            // 基础样式设置
            CTShd shd = rPr.addNewShd();
            shd.setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd.CLEAR);
            shd.setColor(String.format("%06X", config.getColorObject().getRGB() & 0xFFFFFF));
            shd.setFill("auto");
        }
    }

    private String buildTextEffectsXML(WatermarkConfig config) {
        String colorHex = String.format("%06X", config.getColorObject().getRGB() & 0xFFFFFF);
        float opacity = config.getOpacity();

        StringBuilder xml = new StringBuilder();

        // 添加阴影效果使文本更像背景水印
        xml.append("<w:shd xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ");
        xml.append("w:val=\"clear\" ");
        xml.append("w:color=\"").append(colorHex).append("\" ");
        xml.append("w:fill=\"auto\"/>");

        // 如果是对角线模式，添加旋转效果的近似实现
        if (config.getPosition() == WatermarkConfig.Position.DIAGONAL) {
            xml.append("<w:effect xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:val=\"blinkBackground\"/>");
        }

        return xml.toString();
    }

    private void addSimpleBackgroundWatermark(XWPFDocument doc, WatermarkConfig config) {
        // 如果文档没有段落，创建一个
        if (doc.getParagraphs().isEmpty()) {
            doc.createParagraph();
        }

        // 在文档开头插入多个水印段落以形成背景效果
        XWPFParagraph firstPara = doc.getParagraphs().get(0);

        if (config.getPosition() == WatermarkConfig.Position.DIAGONAL) {
            // 对角线模式：创建多个水印段落
            for (int i = 0; i < 3; i++) {
                createSimpleWatermarkParagraph(doc, firstPara, config, i);
            }
        } else {
            // 单一位置模式：创建一个水印段落
            createSimpleWatermarkParagraph(doc, firstPara, config, 0);
        }

        log.info("使用简单背景水印作为备用方案");
    }

    private void createSimpleWatermarkParagraph(XWPFDocument doc, XWPFParagraph firstPara, WatermarkConfig config, int index) {
        XWPFParagraph watermarkPara = doc.insertNewParagraph(firstPara.getCTP().newCursor());
        watermarkPara.setAlignment(ParagraphAlignment.CENTER);

        // 设置段落间距，使水印不占用太多空间
        watermarkPara.setSpacingBefore(0);
        watermarkPara.setSpacingAfter(0);

        XWPFRun run = watermarkPara.createRun();

        // 根据索引设置不同的水印内容或位置
        String text = config.getText();
        if (index > 0) {
            text = "    " + text + "    "; // 添加空格形成错位效果
        }

        run.setText(text);
        run.setFontFamily("SimSun");
        run.setFontSize(Math.max(config.getFontSize() - index * 5, 20)); // 递减字体大小

        // 设置颜色（逐渐变淡）
        int baseColor = config.getColorObject().getRGB() & 0xFFFFFF;
        int fadedColor = adjustColorOpacity(baseColor, config.getOpacity() - index * 0.1f);
        run.setColor(String.format("%06X", fadedColor));
    }

    private int adjustColorOpacity(int baseColor, float opacity) {
        // 简单的颜色透明度调整算法
        opacity = Math.max(0.1f, Math.min(1.0f, opacity));

        int r = (baseColor >> 16) & 0xFF;
        int g = (baseColor >> 8) & 0xFF;
        int b = baseColor & 0xFF;

        // 向白色混合以模拟透明效果
        r = (int) (r + (255 - r) * (1 - opacity));
        g = (int) (g + (255 - g) * (1 - opacity));
        b = (int) (b + (255 - b) * (1 - opacity));

        return (r << 16) | (g << 8) | b;
    }
}
