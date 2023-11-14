package com.csg.supervise;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

/**
 * 附件6
 */
public class Atachment3 {


    private void initData1(String sourceFile,String targetFile) throws Exception {
        Workbook workbook2 = null;
        FileOutputStream fos =null;
        try {
            FileInputStream fileInputStream2 = new FileInputStream(targetFile);
            workbook2 = new XSSFWorkbook(fileInputStream2);
            Map<String,String[]> tableMap=ExcelUtil.getAttachment8TableInfo(sourceFile,1,2);
            List<String[]> dataList=ExcelUtil.getExcelData(sourceFile,2,1);
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
                String tableName=tabs[4];//表中文名
                String busName=tabs[2];//業務名稱
                String a_model=tabs[7];//一級功能
                String b_model=tabs[8];
                String c_model=tabs[9];
                String d_model=tabs[10];

                Row workbook2_sheet1_row = workbook2_sheet1.createRow(i+3);
                workbook2_sheet1_row.createCell(0).setCellValue((i+1));//A Column:序号
                workbook2_sheet1_row.createCell(1).setCellValue(Main.SYSTEM_NAME);//B Column: 系统名称
                workbook2_sheet1_row.createCell(2).setCellValue("网级部署");//C Column: 系统部署级别
                workbook2_sheet1_row.createCell(3).setCellValue(a_model);//D Column: 一级功能名称
                workbook2_sheet1_row.createCell(4).setCellValue(b_model);//E Column: 二级功能名称
                workbook2_sheet1_row.createCell(5).setCellValue(c_model);//F Column:  三级功能名称
                workbook2_sheet1_row.createCell(6).setCellValue(d_model);//G Column:  四级功能名称
                workbook2_sheet1_row.createCell(7).setCellValue("/");//H Column: 页面业务元素统称（非必填）
                workbook2_sheet1_row.createCell(8).setCellValue("/");//I Column: 功能页面录入项
                workbook2_sheet1_row.createCell(9).setCellValue("系统生成");//J Column: 数据来源（选择）

                workbook2_sheet1_row.createCell(10).setCellValue(busName);//K Column: 业务对象名称
                workbook2_sheet1_row.createCell(11).setCellValue(row2);//：L Column:  对应数据表
                workbook2_sheet1_row.createCell(12).setCellValue(tableName);//：M Column:  数据表中文名
                workbook2_sheet1_row.createCell(13).setCellValue(row3);//：N Column:  对应数据字段
                workbook2_sheet1_row.createCell(14).setCellValue(row4);//：O Column:  数据字段中文名
                workbook2_sheet1_row.createCell(15).setCellValue("基础数据");//：P Column:  数据类型
                workbook2_sheet1_row.createCell(16).setCellValue("/");//：Q Column:  统计口径
                workbook2_sheet1_row.createCell(17).setCellValue("CSG-DB-DQ000001");//：R Column:  数据质量标准编号
                workbook2_sheet1_row.createCell(18).setCellValue("/");//：S Column:  完整性
                String value19="其他类:"+row6;

                if (StringUtil.isNotBlank(row4) && row4.contains("编码")) {
                    if (row4.equals("省编码")){
                        value19 = "编码类：附件/信息分类与编码（最新）.zip/《南方电网公司信息分类和编码标准 第5分册 人力资源管理类信息分类和编码.doc》中5.2.组织机构编码\n" +
                                "南方电网公司及二级单位编码\n" +
                                "000000=中国南方电网有限责任公司\n" +
                                "010000=中国南方电网有限责任公司超高压输电公司 \n" +
                                "020000=中国南方电网有限责任公司调峰调频发电公司 ";
                    }else if (row4.equals("局编码")){
                        value19 = "编码类：附件/信息分类与编码（最新）.zip/《南方电网公司信息分类和编码标准 第5分册 人力资源管理类信息分类和编码.doc》中5.2.组织机构编码 中 三级单位编码\n" +
                                "30600=佛山供电局\n" +
                                "31900=东莞供电局";
                    }else{
                        value19 = "编码类:";
                    }

                } else if ((StringUtil.isNotBlank(row4) && (
                        row4.contains("类别")
                                || row4.contains("类型")
                                || row4.contains("分类")
                                || row4.contains("方式")
                                || row4.contains("状态")
                                || row4.contains("标志")
                                || row4.contains("等级")
                                || row4.contains("是否")))) {
                    value19 = "代码类:";
                }else if (!StringUtil.isBlank(row7)) {
                        value19 = "其他类:" + row6 + "(" + row7 + ")";
                }
                workbook2_sheet1_row.createCell(19).setCellValue(value19);//：T Column:  规范性
                workbook2_sheet1_row.createCell(20).setCellValue("/");//：U Column:  一致性
                workbook2_sheet1_row.createCell(21).setCellValue("/");//：V Column:  及时性
                workbook2_sheet1_row.createCell(22).setCellValue("/");//：W Column:  准确性
                workbook2_sheet1_row.createCell(23).setCellValue("/");//：X Column:  依据文件
                workbook2_sheet1_row.createCell(24).setCellValue("/");//：Y Column:  依据条目（页码）
                workbook2_sheet1_row.createCell(25).setCellValue(Main.MANAGER);//：Z Column:  业务归口方负责人
                workbook2_sheet1_row.createCell(26).setCellValue(Main.MANAGER_4A);//：AA Column:  4A账号
                workbook2_sheet1_row.createCell(27).setCellValue(Main.DEVELOPER);//：AB Column:  开发负责人
                workbook2_sheet1_row.createCell(28).setCellValue(Main.DEVELOPER_PHONE);//：AC Column:  联系电话
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
    }
}
