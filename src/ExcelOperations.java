

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperations {

	public static void main(String[] args) throws IOException {
		// read the entire excel sheet
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\testdata2.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet sheet=wb.getSheet("login");
         
        int rowcount=sheet.getLastRowNum()-sheet.getFirstRowNum();
        int cellcount=sheet.getRow(0).getLastCellNum();
    	System.out.println(cellcount);
        
        System.out.println(rowcount);
        for(int i=1;i<=rowcount;i++)
        {
        	
        	 System.out.println();
        	 for(int j=0;j<cellcount;j++)
        	{
        		 
        		System.out.print(sheet.getRow(i).getCell(j).getStringCellValue()+",");
        	}
        	
        }
        //write the data into existing cell in the existed row
        sheet.getRow(4).getCell(0).setCellValue("User4");
        FileOutputStream fout=new FileOutputStream(System.getProperty("user.dir")+"\\src\\testdata2.xlsx");
        wb.write(fout);
        wb.close();
        fis.close();
        fout.close();
       
	}

}
