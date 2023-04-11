package com.csg.supervise;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

/**
 * 附件2
 */
public class Atachment2{

    private void initData1(String sourceFile,String targetFile) throws Exception {
        Workbook workbook2 = null;
        FileOutputStream fos =null;
        try {
            FileInputStream fileInputStream2 = new FileInputStream(targetFile);
            workbook2 = new XSSFWorkbook(fileInputStream2);
            Map<String,String[]> tableMap=ExcelUtil.getAttachment8TableInfo(sourceFile,1,2);
            List<String[]> dataList=ExcelUtil.getExcelData(sourceFile,2,1);
            Sheet workbook2_sheet1 = workbook2.getSheet("DR01业务-数据对象清单");
            fileInputStream2.close();

            //清空行内容
            Integer maxRow =workbook2_sheet1.getLastRowNum();
            for (int i=3;i<=maxRow;i++){
                Row row=workbook2_sheet1.getRow(i);
                if (row==null)
                    continue;
                workbook2_sheet1.removeRow(row);
            }

            maxRow = dataList.size();
            for (int i = 0; i < maxRow; i++) {
                String[]row=dataList.get(i);
               String row0=String.valueOf(row[0]);//A Column: *实例名或TNS
                if (row0.equals("null")||StringUtil.isBlank(row0))
                    continue;
                String row1 = row[2];//B Column: *schema名称/模式名称
                String row2 = row[2];//C Column: *表代码
                String row3 = row[3];//D Column: *字段代码
                String row4 = row[4];;//E Column: *字段名称
                String row5 = row[5];;//F Column: *字段注释
                String row6 = row[6];;//G Column: *字段类型
                String row7 = row[7];;//H Column: *數據長度（没有请填无）

                String[]tabs=tableMap.get(row2);
                String table=tabs[4];//表中文名
                String busName=tabs[2];//業務名稱
                String a_model=tabs[7];//一級功能
                String b_model=tabs[8];
                String c_model=tabs[9];
                String d_model=tabs[10];

                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i+3);
                workbook2_sheet1_row.createCell(0).setCellValue((i+1));//A Column:序号
                workbook2_sheet1_row.createCell(1).setCellValue("数字化");//B Column: 数字化
                workbook2_sheet1_row.createCell(2).setCellValue(busName);//C Column: 业务对象名称
                workbook2_sheet1_row.createCell(3).setCellValue(row4);//D Column: 业务对象信息项名称
                workbook2_sheet1_row.createCell(4).setCellValue("/");//E Column: 所在系统页面截图
                workbook2_sheet1_row.createCell(5).setCellValue("督查督办系统");//F Column: 系统名称
                workbook2_sheet1_row.createCell(6).setCellValue("网级部署");//G Column: 系统部署级别
                workbook2_sheet1_row.createCell(7).setCellValue(a_model);//H Column: 一级功能名称
                workbook2_sheet1_row.createCell(8).setCellValue(b_model);//I Column:  二级功能名称
                workbook2_sheet1_row.createCell(9).setCellValue(c_model);//J Column:  三级功能名称
                workbook2_sheet1_row.createCell(10).setCellValue(d_model);//K Column:  四级功能名称
                workbook2_sheet1_row.createCell(11).setCellValue(row2);//：L Column:  信息项数据表
                workbook2_sheet1_row.createCell(12).setCellValue(table);//：M Column:  信息项数据表中文名
                workbook2_sheet1_row.createCell(13).setCellValue(row3);//：N Column:  信息项数据字段
                workbook2_sheet1_row.createCell(14).setCellValue(row4);//：O Column:  信息项数据字段中文名
                workbook2_sheet1_row.createCell(15).setCellValue(".FD_CREATOR_ID\n" + ".FD_CREATE_TIME");//：P Column:  数据责任人标识
                workbook2_sheet1_row.createCell(16).setCellValue("系统生成");//：Q Column:  数据来源
            }
            fos = new FileOutputStream(targetFile);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (workbook2!=null&&fos!=null)
                workbook2.write(fos);
            if (fos!=null){
                fos.flush();
                fos.close();
            }
        }
    }


    private void initData2(String sourceFile,String targetFile) throws Exception {
        Workbook workbook2 = null;
        FileOutputStream fos =null;
        try {
            FileInputStream fileInputStream2 = new FileInputStream(targetFile);
            workbook2 = new XSSFWorkbook(fileInputStream2);
            Map<String,String[]> tableMap=ExcelUtil.getAttachment8TableInfo(sourceFile,1,2);
            List<String[]> dataList=ExcelUtil.getExcelData(sourceFile,2,1);
            Sheet workbook2_sheet1 = workbook2.getSheet("DR02对象认责矩阵");
            fileInputStream2.close();

            //清空行内容
            Integer maxRow =workbook2_sheet1.getLastRowNum();
            for (int i=1;i<=maxRow;i++){
                Row row=workbook2_sheet1.getRow(i);
                if (row==null)
                    continue;
                workbook2_sheet1.removeRow(row);
            }

            maxRow = dataList.size();
            for (int i = 0; i < maxRow; i++) {
                String[]row=dataList.get(i);
                String row0=String.valueOf(row[0]);//A Column: *实例名或TNS
                if (row0.equals("null")||StringUtil.isBlank(row0))
                    continue;
                String row1 = row[2];//B Column: *schema名称/模式名称
                String row2 = row[2];//C Column: *表代码
                String row3 = row[3];//D Column: *字段代码
                String row4 = row[4];;//E Column: *字段名称
                String row5 = row[5];;//F Column: *字段注释
                String row6 = row[6];;//G Column: *字段类型
                String row7 = row[7];;//H Column: *數據長度（没有请填无）

                String[]tabs=tableMap.get(row2);
                String table=tabs[4];//表中文名
                String busName=tabs[2];//業務名稱
                String a_model=tabs[7];//一級功能
                String b_model=tabs[8];
                String c_model=tabs[9];
                String d_model=tabs[10];

                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i+1);
                workbook2_sheet1_row.createCell(0).setCellValue((i+1));//A Column:序号
                workbook2_sheet1_row.createCell(1).setCellValue("数字化");//B Column: 数字化
                workbook2_sheet1_row.createCell(2).setCellValue(busName);//C Column: 业务对象名称
                workbook2_sheet1_row.createCell(3).setCellValue(row4);//D Column: 业务对象信息项名称


                workbook2_sheet1_row.createCell(4).setCellValue("督查督办系统");//E Column: 系统名称
                workbook2_sheet1_row.createCell(5).setCellValue("网级");//F Column: 单位级别
                workbook2_sheet1_row.createCell(6).setCellValue("南方电网数字电网集团有限公司/南方电网数字企业科技（广东）有限公司/行政人资事业部/主管");//G Column: 业务归口方/岗位
                workbook2_sheet1_row.createCell(7).setCellValue("南方电网数字电网集团有限公司/南方电网数字企业科技（广东）有限公司/行政人资事业部/主管/陈维汉(chenwh3@csg.cn)");//H Column: 业务归口方/岗位/负责人（4A账号）
                workbook2_sheet1_row.createCell(8).setCellValue("南方电网/数字化部/大数据科/数据管理专责");//I Column:  数据管控方/岗位
                workbook2_sheet1_row.createCell(9).setCellValue("数据库中记录的数据创建人");//J Column:  数据录入方/岗位
            }
            fos = new FileOutputStream(targetFile);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (workbook2!=null&&fos!=null)
                workbook2.write(fos);
            if (fos!=null){
                fos.flush();
                fos.close();
            }
        }
    }

    public void initData(String fileName1, String fileName2) throws Exception {
        initData1(fileName1,fileName2);
        initData2(fileName1,fileName2);
    }
}
