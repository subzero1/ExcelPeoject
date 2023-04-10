package com.csg.supervise;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * 附件6
 */
public class Atachment3 {

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
            Sheet workbook2_sheet1 = workbook2.getSheet("DQ01数据质量标准清单");
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
                Cell workbook1_sheet0_row_cell5 = workbook1_sheet1_row.getCell(5);//F Column: Data Type
//                Cell workbook1_sheet0_row_cell6 = workbook1_sheet1_row.getCell(6);//G Column: Length

                String table=workbook1_sheet0_row_cell1.getStringCellValue();
                String busName=String.valueOf(Atachment1.tablesMap.get(table));

                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i+3);
                workbook2_sheet1_row.createCell(0).setCellValue((i+1));//A Column:序号
                workbook2_sheet1_row.createCell(1).setCellValue("督查督办系统");//B Column: 系统名称
                workbook2_sheet1_row.createCell(2).setCellValue("网级部署");//C Column: 系统部署级别
                workbook2_sheet1_row.createCell(3).setCellValue("");//D Column: 一级功能名称
                workbook2_sheet1_row.createCell(4).setCellValue("");//E Column: 二级功能名称
                workbook2_sheet1_row.createCell(5).setCellValue("");//F Column:  三级功能名称
                workbook2_sheet1_row.createCell(6).setCellValue("");//G Column:  四级功能名称
                workbook2_sheet1_row.createCell(7).setCellValue("/");//H Column: 页面业务元素统称（非必填）
                workbook2_sheet1_row.createCell(8).setCellValue("/");//I Column: 功能页面录入项
                workbook2_sheet1_row.createCell(9).setCellValue("督查督办系统");//J Column: 数据来源（选择）

                workbook2_sheet1_row.createCell(10).setCellValue("");//K Column: 业务对象名称
                workbook2_sheet1_row.createCell(11).setCellValue(table);//：L Column:  对应数据表
                workbook2_sheet1_row.createCell(12).setCellValue(busName);//：M Column:  数据表中文名
                workbook2_sheet1_row.createCell(13).setCellValue(workbook1_sheet0_row_cell2.getStringCellValue());//：N Column:  对应数据字段
                workbook2_sheet1_row.createCell(14).setCellValue(workbook1_sheet0_row_cell3.getStringCellValue());//：O Column:  数据字段中文名
                workbook2_sheet1_row.createCell(15).setCellValue("基础数据");//：P Column:  数据类型
                workbook2_sheet1_row.createCell(16).setCellValue("/");//：Q Column:  统计口径
                workbook2_sheet1_row.createCell(17).setCellValue("CSG-DB-DQ000001");//：R Column:  数据质量标准编号
                workbook2_sheet1_row.createCell(18).setCellValue("/");//：S Column:  完整性
                workbook2_sheet1_row.createCell(19).setCellValue("其他类:"+workbook1_sheet0_row_cell5.getStringCellValue());//：T Column:  规范性
                workbook2_sheet1_row.createCell(20).setCellValue("/");//：U Column:  一致性
                workbook2_sheet1_row.createCell(21).setCellValue("/");//：V Column:  及时性
                workbook2_sheet1_row.createCell(22).setCellValue("/");//：W Column:  准确性
                workbook2_sheet1_row.createCell(23).setCellValue("/");//：X Column:  依据文件
                workbook2_sheet1_row.createCell(24).setCellValue("/");//：Y Column:  依据条目（页码）
                workbook2_sheet1_row.createCell(25).setCellValue("陈维汉");//：Z Column:  业务归口方负责人
                workbook2_sheet1_row.createCell(26).setCellValue("chenwh3@csg.cn");//：AA Column:  4A账号
                workbook2_sheet1_row.createCell(27).setCellValue("陆岸亮");//：AB Column:  开发负责人
                workbook2_sheet1_row.createCell(28).setCellValue("18613039080");//：AC Column:  联系电话
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
    }
}
