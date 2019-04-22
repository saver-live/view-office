package com.view.office.controller;

import com.alibaba.fastjson.JSON;
import com.jacob.com.ComThread;
import com.view.office.Config.AppConfig;
import com.view.office.Config.Client;
import com.view.office.commons.R;
import com.view.office.commons.jacob.JacobExcelTool;
import com.view.office.commons.jacob.JacobPptTool;
import com.view.office.commons.jacob.JacobWordTool;
import com.view.office.commons.utils.ClientTool;
import com.view.office.commons.utils.DownloadTool;
import com.view.office.commons.utils.FileUtil;
import com.view.office.commons.excel.ExcelSheet;
import com.view.office.commons.excel.ExcelUtil;
import org.apache.commons.io.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLDecoder;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.UUID;

@RequestMapping("/preview")
@Controller
public class PreviewController {

    @Autowired
    private AppConfig appConfig;

    private static String[] DIRECT_OUT_SUFFICE={"txt","png","jpg","jpeg"};

    @GetMapping("/temp/{path}/{file}")
    public void temp(HttpServletResponse response, @PathVariable("path") String path, @PathVariable("file") String file) {
        File input = new File(appConfig.getFilePath() + path + File.separator + file);
        try {
            FileInputStream fileInputStream = new FileInputStream(input);
            ServletOutputStream outputStream = response.getOutputStream();
            IOUtils.copy(fileInputStream, outputStream);
            fileInputStream.close();
            outputStream.close();
        } catch (IOException e) {
            //文件丢失
        }
    }

    @GetMapping("/error")
    public String error() {
        return "error";
    }

    @GetMapping("/office")
    public void office(@RequestParam("src") String src, HttpServletRequest request, HttpServletResponse response) {
        Boolean registry = Client.getInstance().getRegistry();
        if (registry == null || !registry) {
            Long space = new Date().getTime() - Client.getInstance().getStartDate().getTime();
            if ((space / 1000) > (60 * 60)) {
                try {
                    response.sendRedirect(request.getContextPath() + "/preview/error");
                    return;
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        try {
            String viewPath=null;
            String uuid = UUID.randomUUID().toString().replace("-", "");
            String decode = URLDecoder.decode(src, "utf-8");

            String allow = appConfig.getAllow();
            String[] split = allow.split(";");
            boolean b = ClientTool.allowRequest(split, decode);
            if (!b) {
                response.sendRedirect(request.getContextPath() + "/preview/error");
                return;
            }
            R r = DownloadTool.downLoadFromUrl(decode, appConfig.getFilePath());
            String path = r.getPath();
            String substring = path.substring(path.lastIndexOf(".") + 1);
            List<String> strings = Arrays.asList(DIRECT_OUT_SUFFICE);
            if (strings.contains(substring)){
                viewPath=path;
            }else if (substring.equalsIgnoreCase("doc") || substring.equalsIgnoreCase("docx")) {
                /**
                 * word
                 */
                ComThread.InitMTA(true);
                JacobWordTool jw = new JacobWordTool(false);
                jw.openDocument(path);
                viewPath = appConfig.getFilePath() + uuid + "." + "pdf";
                jw.saveAsPdf(viewPath);
                jw.closeDocument();
                jw.close();
                ComThread.Release();
            } else if (substring.equalsIgnoreCase("xls") || substring.equalsIgnoreCase("xlsx")) {
                /**
                 * excel
                 */
                if (appConfig.getExcel_model() == 1) {
                    ComThread.InitMTA(true);
                    JacobExcelTool je = new JacobExcelTool(false);
                    je.OpenExcel(path, true);
                    viewPath = appConfig.getFilePath() + uuid + "." + "htm";
                    je.saveHtml(viewPath);
                    je.CloseExcel(true);
                    ComThread.Release();
                } else {
                    viewPath = path;
                }
            } else if (substring.equalsIgnoreCase("ppt") || substring.equalsIgnoreCase("pptx")) {
                /**
                 * ppt
                 */
                JacobPptTool jp = new JacobPptTool();
                viewPath = appConfig.getFilePath() + uuid + "." + "pdf";
                jp.saveAsPdf(path, viewPath);
            } else {
                /**
                 * 不支持文件格式
                 */
                response.sendRedirect(request.getContextPath() + "/preview/error");
                return;
            }
            File file = new File(viewPath);

            if (file.exists()) {
                if (viewPath.endsWith(".png")) {
                    FileUtil.copyPdfSteam(file, response);
                }else if (viewPath.endsWith(".jpg")) {
                    FileUtil.copyPdfSteam(file, response);
                }else if (viewPath.endsWith(".jpeg")) {
                    FileUtil.copyPdfSteam(file, response);
                }else if (viewPath.endsWith(".txt")) {
                    FileUtil.copyPdfSteam(file, response);
                }else if (viewPath.endsWith(".pdf")) {
                    FileUtil.copyPdfSteam(file, response);
                } else if (viewPath.endsWith(".htm")) {
                    FileUtil.copyHtmlSteam(file, response, uuid);
                } else if (viewPath.endsWith(".xls") || viewPath.endsWith(".xlsx")) {

                    try {
                        InputStream inputStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("preview/excel.html");
                        BufferedInputStream io = new BufferedInputStream(inputStream);
                        String s = IOUtils.toString(io, "utf-8");
                        List<ExcelSheet> excel;

                        ExcelUtil util = new ExcelUtil();
                        excel = util.getExcelData(viewPath);

                        String replace1 = s.replace("{name}", r.getFileName());
                        String replace = replace1.replace("{excelData}", JSON.toJSONString(excel));
                        FileUtil.copySteam(replace, response);
                        return;
                    } catch (Exception e) {
                        System.out.println(e.toString() + "读取excel模板失败");
                    }
                }
            } else {
                //预览文件不存在,预览失败失败
                //不支持文件格式
                System.out.println("不支持文件格式");
                response.sendRedirect(request.getContextPath() + "/preview/error");
                return;
            }
        } catch (Exception e) {
            e.printStackTrace();
            try {
                System.out.println("异常");
                response.sendRedirect(request.getContextPath() + "/preview/error");
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }

    }
}
