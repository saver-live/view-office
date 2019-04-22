package com.view.office.commands;

import com.view.office.Config.AppConfig;
import com.view.office.Config.Client;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.shell.standard.ShellComponent;
import org.springframework.shell.standard.ShellMethod;

@ShellComponent
public class Commands {

    @Autowired
    private AppConfig appConfig;

    @ShellMethod("查看excel模式：1:html:2:pdf")
    public String setexcel(int param1) {
        appConfig.setExcel_model(param1);
        return "setexcel ok";
    }

    @ShellMethod("激活 软件")
    public String active(String pwd) {
        if (pwd.equals("q123456")) {
            Client.getInstance().setRegistry(true);
            return "active success";
        }
        return "active fail";
    }
}
