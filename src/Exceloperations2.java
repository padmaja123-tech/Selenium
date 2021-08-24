import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exceloperations2 {
	public static void main(String[] args) throws IOException {
		// read the entire excel sheet
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir")+"\\src\\data.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet sheet=wb.getSheet("people");
         
        int rowcount=sheet.getLastRowNum()-sheet.getFirstRowNum();
        System.out.println(rowcount);
       	int cellcount=sheet.getRow(0).getLastCellNum();
       	System.out.println(cellcount);
       	int cellindex=0;

   		for(int j=0;j<cellcount;j++)
   		{	
   		  if(sheet.getRow(0).getCell(j).getStringCellValue().trim().equalsIgnoreCase("age"))
   			  cellindex=j;
   		} 
   		System.out.println(cellindex);
       	for(int i=1;i<=rowcount;i++)
       	{
       		      			  
       		if(sheet.getRow(i).getCell(cellindex).getNumericCellValue()<18)	
              sheet.getRow(i).getCell(cellindex+1).setCellValue("Miner");
            else 
            	 sheet.getRow(i).getCell(cellindex+1).setCellValue("Major");
       			
       		
       	}
        FileOutputStream fout=new FileOutputStream(System.getProperty("user.dir")+"\\src\\data.xlsx");
        wb.write(fout);
        wb.close();
        fis.close();
        fout.close();
       
	}

}
