package Automation.ExcelDriven;

import java.io.IOException;
import java.util.ArrayList;

public class ExcelTest {

	public static void main(String[] args) throws IOException {
		
		excelextract e= new excelextract();
		
		ArrayList some=e.excel("Harshit");
		
		System.out.println(some.get(0));
		System.out.println(some.get(1));
		System.out.println(some.get(2));


		
		



	}

}
