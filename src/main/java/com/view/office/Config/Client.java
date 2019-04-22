package com.view.office.Config;

import java.util.Date;

public class Client {

    private static Client client;

    private Boolean registry;

    private Date startDate;

    private Client() {
    }

    public Boolean getRegistry() {
        return registry;
    }

    public void setRegistry(Boolean registry) {
        this.registry = registry;
    }

    public static Client getInstance() {
        if (client == null) {
            client = new Client();
        }
        return client;
    }

    public Date getStartDate() {
        return startDate;
    }

    public void setStartDate(Date startDate) {
        this.startDate = startDate;
    }
}
