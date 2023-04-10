package com.csg.supervise;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
public class Atachment2{

    private void initData1(String sourceFile,String targetFile) throws Exception {
        Workbook workbook1 = null;
        Workbook workbook2 = null;
        FileOutputStream fos =null;
        try {
            FileInputStream fileInputStream1 = new FileInputStream(sourceFile);
            FileInputStream fileInputStream2 = new FileInputStream(targetFile);
            workbook1 = new XSSFWorkbook(fileInputStream1);
            workbook2 = new XSSFWorkbook(fileInputStream2);
            Sheet workbook1_sheet0 = workbook1.getSheet("Columns");
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

            maxRow = workbook1_sheet0.getLastRowNum();
            for (int i = 0; i < maxRow; i++) {
                Row workbook1_sheet1_row = workbook1_sheet0.getRow(i+1);
                Cell workbook1_sheet0_row_cell0 = workbook1_sheet1_row.getCell(0);//A Column: Owner
                if (workbook1_sheet0_row_cell0 == null || StringUtil.isBlank(workbook1_sheet0_row_cell0.getStringCellValue()))
                    continue;
                Cell workbook1_sheet0_row_cell1 = workbook1_sheet1_row.getCell(1);//B Column: Table
                Cell workbook1_sheet0_row_cell2 = workbook1_sheet1_row.getCell(2);//C Column: Code
                Cell workbook1_sheet0_row_cell3 = workbook1_sheet1_row.getCell(3);//D Column: Name
//                Cell workbook1_sheet0_row_cell4 = workbook1_sheet1_row.getCell(4);//E Column: Comment
//                Cell workbook1_sheet0_row_cell5 = workbook1_sheet1_row.getCell(5);//F Column: Data Type
//                Cell workbook1_sheet0_row_cell6 = workbook1_sheet1_row.getCell(6);//G Column: Length

                String table=workbook1_sheet0_row_cell1.getStringCellValue();
                String busName=String.valueOf(Atachment1.tablesMap.get(table));

                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i+3);
                workbook2_sheet1_row.createCell(0).setCellValue((i+1));//A Column:序号
                workbook2_sheet1_row.createCell(1).setCellValue("数字化");//B Column: 数字化
                workbook2_sheet1_row.createCell(2).setCellValue(busName);//C Column: 业务对象名称
                workbook2_sheet1_row.createCell(3).setCellValue(workbook1_sheet0_row_cell3.getStringCellValue());//D Column: 业务对象信息项名称
                workbook2_sheet1_row.createCell(4).setCellValue("/");//E Column: 所在系统页面截图
                workbook2_sheet1_row.createCell(5).setCellValue("督查督办系统");//F Column: 系统名称
                workbook2_sheet1_row.createCell(6).setCellValue("网级部署");//G Column: 系统部署级别
//                workbook2_sheet1_row.createCell(7).setCellValue("");//H Column: 一级功能名称
//                workbook2_sheet1_row.createCell(8).setCellValue("");//I Column:  二级功能名称
//                workbook2_sheet1_row.createCell(9).setCellValue("");//J Column:  三级功能名称
                workbook2_sheet1_row.createCell(10).setCellValue("/");//K Column:  四级功能名称
                workbook2_sheet1_row.createCell(11).setCellValue(table);//：L Column:  信息项数据表
                workbook2_sheet1_row.createCell(12).setCellValue(busName);//：M Column:  信息项数据表中文名
                workbook2_sheet1_row.createCell(13).setCellValue(workbook1_sheet0_row_cell2.getStringCellValue());//：N Column:  信息项数据字段
                workbook2_sheet1_row.createCell(14).setCellValue(workbook1_sheet0_row_cell3.getStringCellValue());//：O Column:  信息项数据字段中文名
                workbook2_sheet1_row.createCell(15).setCellValue(".DOC_CREATOR_ID\n" + ".DOC_CREATE_TIME");//：P Column:  数据责任人标识
                workbook2_sheet1_row.createCell(16).setCellValue("系统生成");//：Q Column:  数据来源
            }
            fos = new FileOutputStream(targetFile);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (workbook1!=null)
                workbook1.close();
            if (workbook2!=null&&fos!=null)
                workbook2.write(fos);
            if (fos!=null){
                fos.flush();
                fos.close();
            }
        }
    }


    private void initData2(String sourceFile,String targetFile) throws Exception {
        Workbook workbook1 = null;
        Workbook workbook2 = null;
        FileOutputStream fos =null;
        try {
            FileInputStream fileInputStream1 = new FileInputStream(sourceFile);
            FileInputStream fileInputStream2 = new FileInputStream(targetFile);
            workbook1 = new XSSFWorkbook(fileInputStream1);
            workbook2 = new XSSFWorkbook(fileInputStream2);
            Sheet workbook1_sheet0 = workbook1.getSheet("Columns");
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

            maxRow = workbook1_sheet0.getLastRowNum();
            for (int i = 0; i < maxRow; i++) {
                Row workbook1_sheet1_row = workbook1_sheet0.getRow(i+1);
                Cell workbook1_sheet0_row_cell0 = workbook1_sheet1_row.getCell(0);//A Column: Owner
                if (workbook1_sheet0_row_cell0 == null || StringUtil.isBlank(workbook1_sheet0_row_cell0.getStringCellValue()))
                    continue;
                Cell workbook1_sheet0_row_cell1 = workbook1_sheet1_row.getCell(1);//B Column: Table
                Cell workbook1_sheet0_row_cell2 = workbook1_sheet1_row.getCell(2);//C Column: Code
                Cell workbook1_sheet0_row_cell3 = workbook1_sheet1_row.getCell(3);//D Column: Name
//                Cell workbook1_sheet0_row_cell4 = workbook1_sheet1_row.getCell(4);//E Column: Comment
//                Cell workbook1_sheet0_row_cell5 = workbook1_sheet1_row.getCell(5);//F Column: Data Type
//                Cell workbook1_sheet0_row_cell6 = workbook1_sheet1_row.getCell(6);//G Column: Length

                String table=workbook1_sheet0_row_cell1.getStringCellValue();
                String busName=String.valueOf(Atachment1.tablesMap.get(table));

                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i+1);
                workbook2_sheet1_row.createCell(0).setCellValue((i+1));//A Column:序号
                workbook2_sheet1_row.createCell(1).setCellValue("数字化");//B Column: 数字化
                workbook2_sheet1_row.createCell(2).setCellValue(busName);//C Column: 业务对象名称
                workbook2_sheet1_row.createCell(3).setCellValue(workbook1_sheet0_row_cell3.getStringCellValue());//D Column: 业务对象信息项名称


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
            if (workbook1!=null)
                workbook1.close();
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
