package com.lb.json.excel.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

public class JsonUtil {

    private static final Logger logger = LoggerFactory.getLogger(JsonUtil.class);

    /**
     * 读取json文件
     * @param fileName json文件名
     * @return 返回json字符串
     */
    public static String readJsonFile(String fileName) {
        File jsonFile = new File(fileName);
        String jsonStr = "";
        logger.info("————开始读取" + jsonFile.getPath() + "文件————");
        try {
            Reader reader = new InputStreamReader(new FileInputStream(jsonFile), "utf-8");
            int ch = 0;
            StringBuffer sb = new StringBuffer();
            while ((ch = reader.read()) != -1) {
                sb.append((char) ch);
            }
            reader.close();
            jsonStr = sb.toString();
            logger.info("————读取" + jsonFile.getPath() + "文件结束!————");
            return jsonStr;
        } catch (Exception e) {
            logger.info("————读取" + jsonFile.getPath() + "文件出现异常，读取失败!————");
            e.printStackTrace();
            return null;
        }
    }
}
