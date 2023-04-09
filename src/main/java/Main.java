import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Main {
    private static String fileStr;
    static {
//        fileStr=Thread.currentThread().getContextClassLoader().getResource("").getPath();
        fileStr="D://";
    }
    public static void main(String[] args) throws Exception {
        Main.attachment1();
    }

    public static void attachment1() throws Exception {
        String fileName1=fileStr+"Sheet1.xlsx";
        String fileName2=fileStr+"7246-督查督办系统20230309-复审1次.xlsx";
        FileInputStream fileInputStream1=new FileInputStream(fileName1);
        FileInputStream fileInputStream2=new FileInputStream(fileName2);
        Workbook workbook1 = new XSSFWorkbook(fileInputStream1);
        Workbook workbook2 = new XSSFWorkbook(fileInputStream2);
        Sheet workbook1_sheet0=workbook1.getSheetAt(0);
        Sheet workbook2_sheet1=workbook2.getSheetAt(1);
        fileInputStream2.close();
        for (int i=2;i<=30;i++){
            Row workbook1_sheet1_row=workbook1_sheet0.getRow(i);
            Row workbook2_sheet1_row=workbook2_sheet1.createRow(i+1);
            Cell workbook1_sheet0_row_cell=workbook1_sheet1_row.getCell(0);
            Cell workbook2_sheet1_row_cell=workbook2_sheet1_row.createCell(0);
            if (workbook1_sheet0_row_cell==null)
                continue;
            String value=workbook1_sheet0_row_cell.getStringCellValue();
            workbook2_sheet1_row_cell.setCellValue(value);
        }
        workbook1.close();
        FileOutputStream fos=new FileOutputStream(new File(fileName2));
        workbook2.write(fos);
        fos.flush();
        fos.close();
    }
    public static void attachment2(){

    }
    public static void attachment3(){

    }
}
