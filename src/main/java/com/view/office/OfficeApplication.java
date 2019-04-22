package com.view.office;

import com.view.office.Config.Client;
import com.view.office.initializing.SetUp;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.Date;

@SpringBootApplication
public class OfficeApplication {

    public static void main(String[] args) {
        Client.getInstance().setStartDate(new Date());
        SetUp tool = new SetUp();
        tool.setUp();
        SpringApplication.run(OfficeApplication.class, args);
    }

}
