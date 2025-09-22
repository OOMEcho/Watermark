package com.watermark.utils;

import java.awt.Font;
import java.io.InputStream;

public class FontUtils {

    private static final String FONT_PATH = "/fonts/simsun.ttf";
    private static final String FALLBACK_FONT_NAME = "SimSun";

    private FontUtils() {
        // 工具类禁止实例化
    }

    /**
     * 获取 SimSun 字体
     *
     * @param size 字体大小
     * @return 加载到的字体实例
     */
    public static Font getSimSunFont(int size) {
        try (InputStream in = FontUtils.class.getResourceAsStream(FONT_PATH)) {
            if (in == null) {
                throw new IllegalStateException("字体文件未找到: " + FONT_PATH);
            }
            Font baseFont = Font.createFont(Font.TRUETYPE_FONT, in);
            return baseFont.deriveFont(Font.PLAIN, (float) size);
        } catch (Exception e) {
            // 加载失败时回退到系统字体
            return new Font(FALLBACK_FONT_NAME, Font.PLAIN, size);
        }
    }
}
