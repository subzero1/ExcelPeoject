package com.csg.supervise;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.sound.midi.Soundbank;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtil {
    public static String cellFormat(Cell cell) {
        String str = "";
        try {
            if (cell == null) {
                return "";
            }
            switch (cell.getCellType().name()) {
                case "STRING":
                    str = cell.getStringCellValue();
                    break;
                case "BOOLEAN":
                    str = String.valueOf(cell.getBooleanCellValue());
                    break;
                case "NUMERIC":
                    // 先看是否是日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 读取日期格式
                        Date date = cell.getDateCellValue();
                        str = new SimpleDateFormat("yyyy-MM-dd").format(date);
                    } else {
                        str = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case "FORMULA":
                    // 读取公式
                    str = cell.getCellFormula().toString();
                    break;
            }
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return str;
    }

    /**
     * @param sourceFile 文件路径+文件名
     * @param sheetNum  sheetName start by zero
     * @param startNum   起始行 start by zero
     * @return
     */
    public static List<String[]> getExcelData(String sourceFile, Integer sheetNum, Integer startNum) {
        Workbook workbook = null;
        FileInputStream fis = null;
        List<String[]> resultList = new ArrayList<>();
        try {
            fis = new FileInputStream(sourceFile);
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(sheetNum);
            Integer maxRow = sheet.getLastRowNum();
            for (int i = startNum; i <=maxRow; i++) {
                Row row = sheet.getRow(i);
                if (isEmptyRow(row))
                    continue;
                Short maxColumn=row.getLastCellNum();
                String[]obj=new String[maxColumn];
                for (int j=0;j<maxColumn;j++){
                    obj[j]=cellFormat(row.getCell(j));
                }
                resultList.add(obj);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null)
                    workbook.close();
                if (fis != null)
                    fis.close();
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
        return resultList;
    }

    //判断单个单元格是否为空
    public static boolean isEmptyCell(Cell cell) {
        if (cell == null || cell.getCellType().equals(CellType.BLANK)) {
            return true;
        }
        return false;
    }

    /**
     * 判断该行是否为空
     *
     * @param row 行对象
     * @return
     */
    public static boolean isEmptyRow(Row row) {
        //行不存在
        if (row == null) {
            return true;
        }
        //第一个列位置
        int firstCellNum = row.getFirstCellNum();
        //最后一列位置
        int lastCellNum = row.getLastCellNum();
        //空列数量
        int nullCellNum = 0;
        for (int c = firstCellNum; c < lastCellNum; c++) {
            Cell cell = row.getCell(c);
            if (null == cell || CellType.BLANK == cell.getCellType()) {
                nullCellNum++;
                continue;
            }
            cell.setCellType(CellType.STRING);
            String cellValue = cell.getStringCellValue().trim();
            if (StringUtil.isBlank(cellValue)) {
                nullCellNum++;
            }
        }
        //所有列都为空
        if (nullCellNum == (lastCellNum - firstCellNum)) {
            return true;
        }
        return false;
    }

    public static Map<String,String[]> getAttachment8TableInfo(String sourceFile,Integer sheetNum,Integer startNum){

        Map<String,String[]>resultMap=new HashMap();
        List<String[]> list=ExcelUtil.getExcelData(sourceFile,sheetNum,startNum);
        list.forEach(obj->{
            String obj3=String.valueOf(obj[3]);
            resultMap.put(obj3,obj);
        });
        return resultMap;
    }

    public static void main(String[] args) {
        Map reutlMap=ExcelUtil.getAttachment8TableInfo("D://7246-督查督办系统20230309-复审1次.xlsx",1,  2);
        System.out.println(reutlMap);
    }
}
