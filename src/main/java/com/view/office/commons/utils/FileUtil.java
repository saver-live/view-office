package com.view.office.commons.utils;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class FileUtil {

    public static void copySteam(String html, HttpServletResponse response) {
        try {
            ServletOutputStream outputStream = response.getOutputStream();
            ByteArrayInputStream inputStream = new ByteArrayInputStream(html.getBytes("utf-8"));
            IOUtils.copy(inputStream, outputStream);
            inputStream.close();
            outputStream.close();
        } catch (IOException e) {
            //复制失败
        }
    }

    public static void copyPdfSteam(File pdf, HttpServletResponse response) {
        try {
            ServletOutputStream outputStream = response.getOutputStream();
            if (pdf.exists()) {
                FileInputStream stream = new FileInputStream(pdf);
                IOUtils.copy(stream, outputStream);
                if (stream != null) {
                    stream.close();
                }
                if (outputStream != null) {
                    outputStream.close();
                }
            }
        } catch (IOException e) {
//            logger.error("文件复制到response失败");
        }
    }

    public static void copyHtmlSteam(File html, HttpServletResponse response, String uuid) {
        try {
            ServletOutputStream outputStream = response.getOutputStream();
            if (html.exists()) {
                String buffer = FileUtils.readFileToString(html, "utf-8");
                String replace = buffer.replace(uuid + ".files/", "/preview/temp/" + uuid + ".files/");
                byte[] bytes = replace.getBytes();
                ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(bytes);
                IOUtils.copy(byteArrayInputStream, outputStream);
                if (byteArrayInputStream != null) {
                    byteArrayInputStream.close();
                }
                if (outputStream != null) {
                    outputStream.close();
                }
            }
        } catch (IOException e) {
//            logger.error("文件复制到response失败");
        }
    }
}
