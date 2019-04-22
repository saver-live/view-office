package com.view.office.commons.utils;

import com.view.office.commons.R;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLDecoder;
import java.util.UUID;

public class DownloadTool {


    public static R downLoadFromUrl(String urlStr, String savePath) throws IOException {
        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        //设置超时间为3秒
        conn.setConnectTimeout(3 * 1000);
        //防止屏蔽程序抓取而返回403错误
        conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");
        String fileName = getFileName(urlStr, conn);
        //得到输入流
        InputStream inputStream = conn.getInputStream();
        //获取自己数组
        byte[] getData = readInputStream(inputStream);

        //文件保存位置
        File saveDir = new File(savePath);
        if (!saveDir.exists()) {
            saveDir.mkdir();
        }
        String uuid = UUID.randomUUID().toString().replace("-", "");
        File file = new File(saveDir + File.separator + uuid + fileName.substring(fileName.lastIndexOf(".")));
        FileOutputStream fos = new FileOutputStream(file);
        fos.write(getData);
        if (fos != null) {
            fos.close();
        }
        if (inputStream != null) {
            inputStream.close();
        }
        R r=new R(200);
        r.setPath(saveDir + File.separator + uuid + fileName.substring(fileName.lastIndexOf(".")));
        r.setFileName(fileName);
        return r;
    }


    /**
     * 从输入流中获取字节数组
     *
     * @param inputStream
     * @return
     * @throws IOException
     */
    public static byte[] readInputStream(InputStream inputStream) throws IOException {
        byte[] buffer = new byte[1024];
        int len = 0;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while ((len = inputStream.read(buffer)) != -1) {
            bos.write(buffer, 0, len);
        }
        bos.close();
        return bos.toByteArray();
    }

    public static String getFileName(String src, HttpURLConnection conn) throws UnsupportedEncodingException {
        String substring = src.substring(src.lastIndexOf("/") + 1, src.length());
        if (substring.contains(".")) {
            return substring;
        }
        String headerField = conn.getHeaderField("Content-Disposition");
        headerField = headerField.substring(headerField.indexOf("attachment;fileName=") + 20, headerField.length());
        return URLDecoder.decode(headerField.replace("%20", "+"), "utf-8");
    }

//    public static void main(String[] args) {
//
//        try {
//            downLoadFromUrl("http://118.89.50.36:9090/sys/file/download?uuid=5fc31d66fb2d435394b110bdd4f67b8e&x-access-token=5a812e2ef7c498e2cf5e66cb195e981b"
//                    , "D:\\web\\");
//        } catch (Exception e) {
//            // TODO: handle exception
//        }
//    }

}
