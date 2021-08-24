

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperations1 {

	public static void main(String[] args) throws IOException {
		// write the data into the new cell in a new row in excel sheet
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\testdata2.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet sheet=wb.getSheet("login");
         
       // int rowcount=sheet.getLastRowNum()-sheet.getFirstRowNum();
        
        //System.out.println(rowcount);
        XSSFRow row=sheet.createRow(5);
        row.createCell(0).setCellValue("User3");
        row.createCell(1).setCellValue("userpass");
        row.createCell(2).setCellValue("14-07-2014");
        row.createCell(3).setCellValue(5);
        row.createCell(4).setCellValue("passed");
        
        
        FileOutputStream fout=new FileOutputStream(System.getProperty("user.dir")+"\\src\\testdata2.xlsx");
        wb.write(fout);
        wb.close();
        fis.close();
        fout.close();
       
	}

}
