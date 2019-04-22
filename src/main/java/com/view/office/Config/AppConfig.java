package com.view.office.Config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@ConfigurationProperties(prefix = "app")
@Component
public class AppConfig {

    private String filePath;

    /**
     * 1 html查看
     * 2 poi获取数据显示
     */
    private Integer excel_model;

    private String allow;

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public Integer getExcel_model() {
        return excel_model;
    }

    public void setExcel_model(Integer excel_model) {
        this.excel_model = excel_model;
    }

    public String getAllow() {
        return allow;
    }

    public void setAllow(String allow) {
        this.allow = allow;
    }

}
