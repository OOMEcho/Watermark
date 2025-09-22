package com.watermark.utils;

import java.awt.Font;
import java.awt.GraphicsEnvironment;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

public class FontUtils {

    private static final String FONT_PATH = "/fonts/simsun.ttf";

    // 常见系统中文字体列表，可根据需要扩展
    private static final List<String> CHINESE_FALLBACK_FONTS = Arrays.asList(
            "SimSun", "Microsoft YaHei", "PingFang SC", "Heiti SC", "WenQuanYi Micro Hei"
    );

    private FontUtils() {
        // 工具类禁止实例化
    }

    /**
     * 获取中文字体
     *
     * @param size 字体大小
     * @return 可用的字体
     */
    public static Font getChineseFont(int size) {
        // 1. 尝试加载自定义字体
        try (InputStream in = FontUtils.class.getResourceAsStream(FONT_PATH)) {
            if (in != null) {
                Font baseFont = Font.createFont(Font.TRUETYPE_FONT, in);
                return baseFont.deriveFont(Font.PLAIN, (float) size);
            }
        } catch (Exception ignored) {
        }

        // 2. 尝试从系统字体中选一个支持中文的
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        List<String> availableFonts = Arrays.asList(ge.getAvailableFontFamilyNames());

        for (String fallback : CHINESE_FALLBACK_FONTS) {
            if (availableFonts.contains(fallback)) {
                return new Font(fallback, Font.PLAIN, size);
            }
        }

        // 3. 如果都没有，则使用逻辑字体保证不报错
        return new Font(Font.SANS_SERIF, Font.PLAIN, size);
    }
}
