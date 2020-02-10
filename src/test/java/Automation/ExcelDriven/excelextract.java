package Automation.ExcelDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class excelextract {
	

	public ArrayList<String> excel(String TestCaseName) throws IOException {
		
		ArrayList ar= new ArrayList();

		// system.getProperty("user.dir")
		FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\ExcelSheet.xlsx");
		XSSFWorkbook Workbook = new XSSFWorkbook(fis);

		int count = Workbook.getNumberOfSheets();
		for (int i = 0; i < count; i++) {

			if (Workbook.getSheetName(i).equals("Sheet1")) {

				XSSFSheet sheet = Workbook.getSheetAt(i);

				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();

				Iterator<Cell> cell = firstrow.cellIterator();

				int k = 0;
				int coloumn = 0;
				while (cell.hasNext()) {

					Cell firstcell = cell.next();
					if (firstcell.getStringCellValue().equalsIgnoreCase("Name")) {

						coloumn = k;

					}
					k++;
				}
				System.out.println(coloumn);
				
				while(rows.hasNext()) {
					
					Row r=rows.next();
					if(r.getCell(coloumn).getStringCellValue().equals(TestCaseName)) {
						
						Iterator<Cell> cs=r.cellIterator();
						
						while(cs.hasNext()) {
						
							Cell val=cs.next();
							if(val.getCellType()==Cell.CELL_TYPE_STRING) {
							
							ar.add(val.getStringCellValue());
							}else{
								
								ar.add(NumberToTextConverter.toText(val.getNumericCellValue()));
								
							}
							
							}
						
					
						
						}
					}
		
				}


		}
	
		
		return ar;
		}
}