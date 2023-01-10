package pkg1;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writexlsx 
{
	public static void main(String[] args) throws IOException
	{
		File f=new File("../ProjectA/Apache poii.xlsx");  // connection establish
		FileOutputStream fo=new FileOutputStream(f); // output stream object
		XSSFWorkbook xs=new XSSFWorkbook();  // workbook object
		XSSFSheet xt=xs.createSheet("SheetA");   // sheet object
		for (int i=0; i<3; i++)   //  loop for row
		{
			 XSSFRow xr=xt.createRow(i);  // row object
			 for (int j=0; j<5; j++)   // loop for column
			 {
				 XSSFCell xc=xr.createCell(j); // column object
				 xc.setCellValue("Dheeraj");  // set cell data
			 }
		}
xs.write(fo); // will move the data from workbook to output stream
fo.flush(); // will move the data from output stream to file
fo.close();  // for saving it
	}
}
