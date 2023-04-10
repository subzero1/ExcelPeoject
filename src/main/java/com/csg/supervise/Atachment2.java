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
    public void initData1(String sourceFile,String targetFile) throws Exception {
        Workbook workbook1 = null;
        Workbook workbook2 = null;
        FileOutputStream fos =null;
        try {
            FileInputStream fileInputStream1 = new FileInputStream(sourceFile);
            FileInputStream fileInputStream2 = new FileInputStream(targetFile);
            workbook2 = new XSSFWorkbook(fileInputStream2);
            workbook1 = new XSSFWorkbook(fileInputStream1);
            Sheet workbook1_sheet0 = workbook1.getSheet("Columns");
            Sheet workbook2_sheet1 = workbook2.getSheet("DR01业务-数据对象清单");
            fileInputStream2.close();

            //清空内容
            Integer maxRow =workbook2_sheet1.getLastRowNum();
            for (int i=3;i<=maxRow;i++){
                Row row=workbook2_sheet1.getRow(i);
                if (row==null)
                    continue;
                workbook2_sheet1.removeRow(row);
            }

            maxRow = workbook1_sheet0.getLastRowNum();
            for (int i = 2; i <= maxRow; i++) {
                Row workbook1_sheet1_row = workbook1_sheet0.getRow(i);
                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i);
                Cell workbook1_sheet0_row_cell0 = workbook1_sheet1_row.getCell(0);//Owner
                Cell workbook1_sheet0_row_cell1 = workbook1_sheet1_row.getCell(1);//Name
                Cell workbook1_sheet0_row_cell2 = workbook1_sheet1_row.getCell(2);//Code
                Cell workbook1_sheet0_row_cell3 = workbook1_sheet1_row.getCell(3);//Comment

                if (workbook1_sheet0_row_cell0 == null || StringUtil.isBlank(workbook1_sheet0_row_cell0.getStringCellValue()))
                    continue;
                workbook2_sheet1_row.createCell(0).setCellValue("EKP");//A Column:*实例名或TNS
                workbook2_sheet1_row.createCell(1).setCellValue(workbook1_sheet0_row_cell0.getStringCellValue().replace("User '", "").replace("'", ""));//B Column: *schema名称/模式名称
                workbook2_sheet1_row.createCell(2).setCellValue(workbook1_sheet0_row_cell1.getStringCellValue());//C Column: *业务对象业务对象
                workbook2_sheet1_row.createCell(3).setCellValue(workbook1_sheet0_row_cell2.getStringCellValue());//D Column: *表代码
                workbook2_sheet1_row.createCell(4).setCellValue(workbook1_sheet0_row_cell1.getStringCellValue());//E Column: *表名称
                workbook2_sheet1_row.createCell(5).setCellValue(workbook1_sheet0_row_cell3.getStringCellValue());//F Column: *表注释
                workbook2_sheet1_row.createCell(6).setCellValue("更改");//G Column:*操作类型
                workbook2_sheet1_row.createCell(7).setCellValue("");//H Column:*一级功能名称
                workbook2_sheet1_row.createCell(8).setCellValue("");//I Column:*二级功能名称
                workbook2_sheet1_row.createCell(9).setCellValue("");//J Column:*三级功能名称
                workbook2_sheet1_row.createCell(10).setCellValue("无");//K Column:*四级功能名称
                workbook2_sheet1_row.createCell(11).setCellValue("运维部门");//L Column:*数据责任部门
                workbook2_sheet1_row.createCell(12).setCellValue("数据岗");//M Column:*数据责任岗位（不清楚的请填无）
                workbook2_sheet1_row.createCell(13).setCellValue("不涉密");//N Column:*密级（不涉密/核心商密/普通商密/工作秘密/敏感信息）
                workbook2_sheet1_row.createCell(14).setCellValue("√");//O Column:总部
                for (int j=15;j<=23;j++){
                    workbook2_sheet1_row.createCell(j).setCellValue("无");
                }
            }
            fos = new FileOutputStream(targetFile);
        }catch (Exception e){
            e.printStackTrace();
        } finally {
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

    public void initData2(String sourceFile,String targetFile) throws Exception {
        Workbook workbook1 = null;
        Workbook workbook2 = null;
        FileOutputStream fos =null;
        String total="";
        try {
            FileInputStream fileInputStream1 = new FileInputStream(sourceFile);
            FileInputStream fileInputStream2 = new FileInputStream(targetFile);
            workbook1 = new XSSFWorkbook(fileInputStream1);
            workbook2 = new XSSFWorkbook(fileInputStream2);
            Sheet workbook1_sheet0 = workbook1.getSheet("Columns");
            Sheet workbook2_sheet1 = workbook2.getSheetAt(2);
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
            for (int i = 1; i <= maxRow; i++) {
                total=""+i;
                Row workbook1_sheet1_row = workbook1_sheet0.getRow(i);
                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i);
                Cell workbook1_sheet0_row_cell0 = workbook1_sheet1_row.getCell(0);//A Column: Owner
                Cell workbook1_sheet0_row_cell1 = workbook1_sheet1_row.getCell(1);//B Column: Table
                Cell workbook1_sheet0_row_cell2 = workbook1_sheet1_row.getCell(2);//C Column: Code
                Cell workbook1_sheet0_row_cell3 = workbook1_sheet1_row.getCell(3);//D Column: Name
                Cell workbook1_sheet0_row_cell4 = workbook1_sheet1_row.getCell(4);//E Column: Comment
                Cell workbook1_sheet0_row_cell5 = workbook1_sheet1_row.getCell(5);//F Column: Data Type
                Cell workbook1_sheet0_row_cell6 = workbook1_sheet1_row.getCell(6);//G Column: Length

                if (workbook1_sheet0_row_cell0 == null || StringUtil.isBlank(workbook1_sheet0_row_cell0.getStringCellValue()))
                    continue;
                workbook2_sheet1_row.createCell(0).setCellValue("EKP");//A Column:*实例名或TNS
                workbook2_sheet1_row.createCell(1).setCellValue(workbook1_sheet0_row_cell0.getStringCellValue().replace("User '", "").replace("'", ""));//B Column: *schema名称/模式名称
                workbook2_sheet1_row.createCell(2).setCellValue(workbook1_sheet0_row_cell1.getStringCellValue());//C Column: *表代码
                workbook2_sheet1_row.createCell(3).setCellValue(workbook1_sheet0_row_cell2.getStringCellValue());//D Column: *字段代码
                workbook2_sheet1_row.createCell(4).setCellValue(workbook1_sheet0_row_cell3.getStringCellValue());//E Column: *字段名称
                workbook2_sheet1_row.createCell(5).setCellValue(workbook1_sheet0_row_cell4.getStringCellValue());//F Column: *字段注释
                String dataType=workbook1_sheet0_row_cell5.getStringCellValue();
                if (dataType.contains("(")){
                    dataType=dataType.substring(0,dataType.indexOf("("));
                }
                workbook2_sheet1_row.createCell(6).setCellValue(dataType);//G Column: *字段类型
                String length=ExcelUtil.cellFormat(workbook1_sheet0_row_cell6) ;
                if (length.equals("0.0")){
                    length="";
                }else if (length.contains(".0")){
                    length=length.replace(".0","");
                }
                workbook2_sheet1_row.createCell(7).setCellValue(length);//H Column: *数据长度（没有请填无）
                workbook2_sheet1_row.createCell(8).setCellValue("更改");//I Column: *操作类型
                workbook2_sheet1_row.createCell(9).setCellValue("无");//J Column: *数据精度（没有请填无）
                workbook2_sheet1_row.createCell(10).setCellValue("是");//K Column: *公司内部是否可共享
                workbook2_sheet1_row.createCell(11).setCellValue("无条件共享");//L Column: *共享类型（无条件共享/有条件共享/不共享）
                workbook2_sheet1_row.createCell(12).setCellValue("无");//M Column: *共享条件（如是有条件共享，无则填无）
                workbook2_sheet1_row.createCell(13).setCellValue("有条件开放");//N Column: *开放类型（无条件开放/有条件开放/不开放）
                workbook2_sheet1_row.createCell(14).setCellValue("通过审核后开放");//O Column: *开放条件（如是有条件开放，无则填无）
                workbook2_sheet1_row.createCell(15).setCellValue("总部");//P Column: *区域权限（以单位为区域范围）
                workbook2_sheet1_row.createCell(16).setCellValue("行政部");//Q Column: *部门权限（数据责任部门）
                workbook2_sheet1_row.createCell(17).setCellValue("公司领导、部门领导、主管、一般用户");//R *角色权限（公司领导、部门领导、主管、一般用户）
                workbook2_sheet1_row.createCell(18).setCellValue("否");//S Column: *共享过程是否脱敏
                workbook2_sheet1_row.createCell(19).setCellValue("无");//T Column: *脱敏要求
                workbook2_sheet1_row.createCell(20).setCellValue("是");//U Column: *满足脱敏要求后是否可对外开放
                workbook2_sheet1_row.createCell(21).setCellValue("业务数据");//V Column: *数据分类（业务数据/客户数据/重要数据）
                workbook2_sheet1_row.createCell(22).setCellValue("一级");//W Column: *安全级别（一级/二级/三级）
            }
            fos = new FileOutputStream(targetFile);
        }catch (Exception e){
            System.out.println(total);
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
