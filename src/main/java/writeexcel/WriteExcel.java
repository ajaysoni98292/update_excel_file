package writeexcel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author ajay
 */
public class WriteExcel {

    public static void writeCountryListToFile(String fileName, String text) throws Exception {
        FileInputStream fsIP= new FileInputStream(new File(fileName));
        HSSFWorkbook wb = new HSSFWorkbook(fsIP);
        HSSFSheet worksheet = wb.getSheetAt(0);
        Cell cell = null;
        cell = worksheet.getRow(1).getCell(1);
        cell.setCellValue(text);
        fsIP.close();
        FileOutputStream output_file =new FileOutputStream(new File(fileName));  //Open FileOutputStream to write updates
        wb.write(output_file); 
        output_file.close();
    }

    public static void main(String args[]) throws Exception {
        WriteExcel.writeCountryListToFile("F:\\ZTrush\\Test.xls", "Ajay");
    }
}
