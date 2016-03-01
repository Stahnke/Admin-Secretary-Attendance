import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToText {

	public static void main (String args[]) throws IOException
	{
		//Output file
		PrintWriter writer = new PrintWriter("Stocks.txt", "UTF-8");
		
		//Input file
		FileInputStream fis = new FileInputStream(args[0]);
		XSSFWorkbook in_book = new XSSFWorkbook(fis);
		XSSFSheet in_sheet = in_book.getSheetAt(0);
		FormulaEvaluator formulaEvaluator = in_book.getCreationHelper().createFormulaEvaluator();
		
		for(Row row : in_sheet)
		{
			Cell cell = row.getCell(0);
			String temp_string = cell.getStringCellValue();
			writer.println(temp_string);
			
		}
		
		writer.close();
	
	}
	
	
}
