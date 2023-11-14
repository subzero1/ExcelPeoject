package com.csg.supervise;

import org.openxmlformats.schemas.drawingml.x2006.chart.STGapAmount;

import java.io.IOException;

public class Main {
    private static String fileStr = "D://";
    private static String fileName1 = fileStr + "Sheet1.xlsx";
    private static String fileName2 = fileStr + "附件8：元数据清单.xlsx";
    private static String fileName3 = fileStr + "附件2：数据认责矩阵.xlsx";
    private static String fileName4 = fileStr + "附件6：数据质量标准与数据质量规则清单.xlsx";

    public static  String SYSTEM_NAME="督查督办系统";
    public static  String DOMAIN_NAME="综合管理域";


    public static String MANAGER="陈维汉";
//    public static String MANAGER_PHONE="13750003015";
    public static String MANAGER_4A="chenwh3@csg.cn";
    public static String DEVELOPER="佘金程";
    public static String DEVELOPER_PHONE="18207160445";

    /**
     * 先执行附件1，再执行附件2,3
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        fileStr = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        System.out.println("=========================================================");
        System.out.println("請輸入執行順序:\n1、生成附件8 \n2、生成附件2,6");
        System.out.println("請輸入數字:");
        char i = (char) System.in.read();
        System.out.println("your char is :" + i);
        String s=String.valueOf(i);
        if (s.equals("1")){
            new Atachment1().initData(fileName1, fileName2);
        }else{
            new Atachment2().initData(fileName2, fileName3);
            new Atachment3().initData(fileName2, fileName4);
        }
    }

}
