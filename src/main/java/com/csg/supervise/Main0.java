package com.csg.supervise;

public class Main0 {
    private static String fileStr = "D://";
    private static String fileName1 = fileStr + "Sheet1.xlsx";
    private static String fileName2 = fileStr + "7246-督查督办系统20230309-复审1次.xlsx";
    public static void main(String[] args) throws Exception {
        Atachment1 atachment1=new Atachment1();
        atachment1.initData(fileName1,fileName2);
    }

}
