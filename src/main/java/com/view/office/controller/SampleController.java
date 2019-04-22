package com.view.office.controller;

import org.apache.commons.io.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

@Controller
@RequestMapping("/sample")
public class SampleController {

    @Autowired
    private PreviewController previewController;

    @GetMapping("/download/{file}")
    public void download(HttpServletResponse response,@PathVariable("file") String file) throws IOException {
        InputStream stream=null;
        String fileName="";
        if (file.equals("word")){
            fileName="word.docx";
            stream = Thread.currentThread().getContextClassLoader().getResourceAsStream("file/word.docx");
        }else if (file.equals("excel")){
            fileName="excel.xlsx";
            stream = Thread.currentThread().getContextClassLoader().getResourceAsStream("file/excel.xlsx");
        }else if (file.equals("ppt")){
            fileName="ppt.pptx";
            stream = Thread.currentThread().getContextClassLoader().getResourceAsStream("file/ppt.pptx");
        }else if (file.equals("txt")){
            fileName="text.txt";
            stream = Thread.currentThread().getContextClassLoader().getResourceAsStream("file/text.txt");
        }
        ServletOutputStream outputStream = response.getOutputStream();
        response.setContentType("multipart/form-data");
        response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(fileName, "utf-8").replace("+", "%20"));
        IOUtils.copy(stream, outputStream);
        stream.close();
        outputStream.close();
    }

    @GetMapping("/preview/{file}")
    public void preview(@PathVariable("file") String file,HttpServletRequest request, HttpServletResponse response){
        String url="http://127.0.0.1/sample/download/"+file;
        try {
            previewController.office(URLEncoder.encode(url,"utf-8"),request,response);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
    }
}
