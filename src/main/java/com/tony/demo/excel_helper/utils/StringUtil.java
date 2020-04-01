package com.tony.demo.excel_helper.utils;

/**
 * Created by TonyZeng on 2017/9/6.
 */
public class StringUtil {
    /**
     * 转换驼峰命名方式
     * @param name
     * @return
     */
    public static String toCapitalizeCamelCase(String name) {
        if (name == null) {
            return null;
        }
        //name = name.toLowerCase();

        StringBuilder sb = new StringBuilder(name.length());
        boolean upperCase = false;
        for (int i = 0; i < name.length(); i++) {
            char c = name.charAt(i);

            if (c == '_') {
                upperCase = true;
            } else if (upperCase) {
                sb.append(Character.toUpperCase(c));
                upperCase = false;
            } else {
                sb.append(c);
            }
        }
        name = sb.toString();
        return name.substring(0, 1).toUpperCase() + name.substring(1);
    }
}
