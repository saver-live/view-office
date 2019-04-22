package com.view.office.initializing;


import org.apache.commons.io.IOUtils;

import java.io.*;
import java.util.Map;

public class SetUp {

    public void setUp() {
        Map<String, String> getenv = System.getenv();
        String os = getenv.get("OS");
        if (!os.startsWith("Win")) {
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");
            System.out.println("不支持当前操作系统。。。。。。................................................。。。。。。。。。");

            return;
        }
        String java_home = getenv.get("JAVA_HOME");
        String SystemDrive = getenv.get("SystemDrive");
        String path = SystemDrive + "\\Windows\\System32";
        copy(path);
        String bin = java_home + "\\bin";
        copy(bin);
        String jre = java_home + "\\jre\\bin";
        copy(jre);
    }

    public void copy(String path) {
        InputStream inputStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("jacob/jacob-1.17-x64.dll");
        File out = new File(path);
        if (!out.exists()) {

            try {
                FileOutputStream fileOutputStream = new FileOutputStream(out);
                IOUtils.copy(inputStream, fileOutputStream);
                fileOutputStream.close();
                inputStream.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
    }
}
