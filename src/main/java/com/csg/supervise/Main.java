package com.csg.supervise;

public class Main {
    private static String fileStr = "D://";
    private static String fileName1 = fileStr + "Sheet1.xlsx";
    private static String fileName2 = fileStr + "7246-督查督办系统20230309-复审1次.xlsx";
    private static String fileName3 = fileStr + "附件2：南方电网公司数据认责矩阵20230309.xlsx";
    public static void main(String[] args) throws Exception {
        Atachment1 atachment1=new Atachment1();
        atachment1.initData(fileName1,fileName2);
        Atachment2 atachment2=new Atachment2();
        atachment2.initData(fileName1,fileName3);
    }

}
