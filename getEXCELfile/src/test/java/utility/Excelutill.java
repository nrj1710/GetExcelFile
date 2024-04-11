package utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class Excelutill {

	public static void main(String[] args) throws Throwable{
		Excelutill obj = new Excelutill();
		
		obj.WriteExcel("Sheet1", 2, 2,"pass" );
		
			
		
     }
	public void WriteExcel(String sheetName ,int rNum ,int cNum,String DATA) throws Throwable {
		
			FileInputStream fis =new FileInputStream("C:\\Users\\niraj\\eclipse-workspace\\getEXCELfile\\inputFile\\Book1.xlsx");
			Workbook wb =WorkbookFactory.create(fis);
			
			Sheet s =wb.getSheet(sheetName);
			Row r =s.getRow(rNum);
			Cell c= r.createCell(cNum);
			
			c.setCellValue(DATA);
		     
			FileOutputStream fos = new FileOutputStream("C:\\\\Users\\\\niraj\\\\eclipse-workspace\\\\getEXCELfile\\\\inputFile\\\\Book1.xlsx");
			wb.write(fos);
//			for(int i= 1; i<=rowCount; i++) {
//				for(int j=2; j<=cellcount; j++) ;
			}
	
	}

