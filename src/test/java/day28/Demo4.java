package day28;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo4 {

	public static void main(String[] args) throws Exception 
	{

		Workbook wb = WorkbookFactory.create(new File("./data/Selenium2.xlsx"));
		int rc = wb.getSheet("sheet2").getLastRowNum();
		
		for(int i=0;i<=rc;i++)
		{
			try {
					int cc=wb.getSheet("sheet2").getRow(i).getLastCellNum();
					for(int j=0;j<cc;j++) 
					{
						try 
						{
							String v = wb.getSheet("sheet2").getRow(i).getCell(j).getStringCellValue();
							System.out.print(v+" ");
						}
						catch (NullPointerException e) 
						{
							System.out.print("-- ");
						}
					}
			}
			catch (NullPointerException e) 
			{
				System.out.print("-- -- --");
			}
			System.out.println();
		}
		wb.close();
		
	}

}
