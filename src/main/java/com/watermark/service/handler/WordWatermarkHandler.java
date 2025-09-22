package com.watermark.service.handler;

import com.watermark.config.WatermarkConfig;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import java.io.InputStream;
import java.io.OutputStream;

public class WordWatermarkHandler {

    public void addWatermark(InputStream input, OutputStream output, WatermarkConfig config) throws Exception {
        XWPFDocument doc = new XWPFDocument(input);

        addBackgroundWatermark(doc, config);

        doc.write(output);
        doc.close();
    }

    private void addBackgroundWatermark(XWPFDocument doc, WatermarkConfig config) {
        XWPFHeaderFooterPolicy policy = doc.getHeaderFooterPolicy();
        if (policy == null) {
            policy = doc.createHeaderFooterPolicy();
        }

        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph watermarkPara = header.createParagraph();
        watermarkPara.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = watermarkPara.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(config.getFontSize());
        run.setText(config.getText());

        String colorHex = String.format("%06X", config.getColor().getRGB() & 0xFFFFFF);
        run.setColor(colorHex);

        CTPPr ppr = watermarkPara.getCTP().getPPr();
        if (ppr == null) {
            ppr = watermarkPara.getCTP().addNewPPr();
        }

        setWatermarkPosition(ppr, config.getPosition());
    }

    private void setWatermarkPosition(CTPPr ppr, WatermarkConfig.Position position) {
        CTJc jc = ppr.getJc();
        if (jc == null) {
            jc = ppr.addNewJc();
        }

        switch (position) {
            case TOP_LEFT:
                jc.setVal(STJc.LEFT);
                break;
            case TOP_RIGHT:
                jc.setVal(STJc.RIGHT);
                break;
            default:
                jc.setVal(STJc.CENTER);
        }
    }
}
