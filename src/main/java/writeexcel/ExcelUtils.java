package writeexcel;

import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author ajay
 */
public class ExcelUtils {

    public static String read(String fileName, int rowNumber, int cellNumber) throws Exception {
        FileInputStream fileInputStream= new FileInputStream(new File(fileName));
        HSSFWorkbook wb = new HSSFWorkbook(fileInputStream);
        HSSFSheet worksheet = wb.getSheetAt(0);
        return worksheet.getRow(rowNumber).getCell(cellNumber).getStringCellValue();
    }

    public static void write(String fileName, int rowNumber, int cellNumber, String text) throws Exception {
        FileInputStream fileInputStream= new FileInputStream(new File(fileName));
        HSSFWorkbook wb = new HSSFWorkbook(fileInputStream);
        HSSFSheet worksheet = wb.getSheetAt(0);
        Cell cell = worksheet.getRow(rowNumber).getCell(cellNumber);
        if (cell == null) {
            Row row = worksheet.createRow(rowNumber);
            cell = row.createCell(cellNumber);
        }
        cell.setCellValue(text);
        fileInputStream.close();
        FileOutputStream fileOutputStream =new FileOutputStream(new File(fileName));
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }

    public static void main(String args[]) throws Exception {
        String path = "F:\\ZTrush\\Test.xls";
        String xpath = ExcelUtils.read(path, 1, 1);
        System.out.println(xpath);
        ExcelUtils.write(path, 2,1,"Ram jane");
    }
}
