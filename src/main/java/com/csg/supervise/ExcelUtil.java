package com.csg.supervise;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelUtil {
    public static String cellFormat(Cell cell) {
        String str = "";
        try {
            if (cell == null) {
                return "";
            }
            switch (cell.getCellType().name()) {
                case "STRING" :
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
}
