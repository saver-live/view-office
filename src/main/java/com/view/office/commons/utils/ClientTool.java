package com.view.office.commons.utils;

public class ClientTool {

    public static boolean allowRequest(String[] arr, String ip) {
        for (String s : arr) {
            if (ip.startsWith(s)) {
                return true;
            }
        }
        return false;
    }
}
