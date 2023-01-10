package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readxlsx
{
	public static void main(String[] args) throws IOException 
	{
		File f=new File("../ProjectA/Apache poi.xlsx"); // connection establish
		FileInputStream fi=new FileInputStream(f);   // input stream object
		XSSFWorkbook xs=new XSSFWorkbook(fi);    // workbook object
		XSSFSheet xt=xs.getSheetAt(0);             // sheet object
		int r=xt.getPhysicalNumberOfRows();   // fetch no. of rows from sheet
		for (int i=0; i<r; i++)  //  loop for row
		{
			XSSFRow xr=xt.getRow(i); // every time it will create row object
			int c=xr.getPhysicalNumberOfCells();  // // fetch no. of columns
			for (int j=0; j<c; j++)   // loop for column
			{
				XSSFCell xc=xr.getCell(j);  // every time it will create column object
				System.out.println(xc.getNumericCellValue());  //fetch the content of column
			}
		}
			}
}
