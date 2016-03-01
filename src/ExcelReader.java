import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import javax.swing.JPanel;
import javax.swing.JTextArea;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	public static void main (String args[]) throws IOException
	{
		
		
	}
	
	public void getExcelSheet(List<File> myFiles) throws IOException
	{
		String GBM = "GBM #";
		String temp_string;
		int temp_code;
		double[][][] attendance_array = new double[5][200][myFiles.size()];
		double[][][] attendance_totals = new double[5][200][1];
		HashMap<Integer, String>[] all_maps = new HashMap[5];
		TreeMap<String, Integer>[] all_sorted_maps = new TreeMap[5];
		String[] family_names = new String[]{"Alpha", "Phi", "Omega", "Rho", "Pi"};
		
		for(int a = 0; a < 5; a++)
		{
			for(int b = 0; b < 200; b++)
			{
				attendance_totals[a][b][0] = 0;
				for(int c = 0; c < myFiles.size(); c++)
				{
					attendance_array[a][b][c] = 0;
				}
			}
		}
		
		//output file setup
		XSSFWorkbook out_book = new XSSFWorkbook();
		XSSFSheet out_sheet = out_book.createSheet("FirstExcelSheet");
		
		//Set up each family's hashmap
		for(int x = 0; x < all_maps.length; x++)
		{
			all_maps[x] = new HashMap<Integer, String>();
		}	
		
		//For each input file, read the file and check attendance
		for(int i = 0; i < myFiles.size(); i++)
		{
			//input file setup
			FileInputStream fis = new FileInputStream((myFiles.get(i)));
			XSSFWorkbook in_book = new XSSFWorkbook(fis);
			XSSFSheet in_sheet = in_book.getSheetAt(0);
			FormulaEvaluator formulaEvaluator = in_book.getCreationHelper().createFormulaEvaluator();
	
			//Traverse each row
			for(Row row : in_sheet)
			{
				//Traverse each column
				for(Cell cell : row){
						
					switch(formulaEvaluator.evaluateInCell(cell).getCellType())
					{
					case Cell.CELL_TYPE_STRING:
						if(row.getRowNum() != 0) //Skip the family names
						{
							//Get the name in the current cell
							temp_string = cell.getStringCellValue();
							//Develop a hashing code for that name
							temp_code = temp_string.hashCode();
							temp_code = Math.abs(temp_code);
							temp_code = temp_code % 200;
							while(all_maps[cell.getColumnIndex()].containsKey(temp_code))
							{
								if(all_maps[cell.getColumnIndex()].get(temp_code).equals(temp_string))
									break;
								temp_code++;
								if(temp_code >= 200)
								temp_code = 0;
							}
							//Store the name at the hashcode for the current family
							all_maps[cell.getColumnIndex()].put(temp_code, temp_string);
							//Store beside the current attendance being marked (Ex: SIGNIN1, SIGNOUT3...)
							attendance_array[cell.getColumnIndex()][temp_code][i] = 0.5;
							attendance_totals[cell.getColumnIndex()][temp_code][0] += 0.5;
								//System.out.print(temp_string + "\t\t");
						}
						break;
					}
				}
			}
			in_book.close();
		}
		
			//Sort each family alphabetically
			for(int y = 0; y < all_maps.length; y++)
			{
				//For the current family
				TreeMap<String, Integer> sorted_names = new TreeMap<String, Integer>();
				//Sort that family
				for (Map.Entry entry : all_maps[y].entrySet()) 
				{
					sorted_names.put((String) entry.getValue(), (Integer)entry.getKey());
				}
				//Store the sorted family
				all_sorted_maps[y] = sorted_names;
			}
			
			//Write the header row (ex: Families, GBM#1, GBM#2....Totals)
			XSSFRow out_row = out_sheet.createRow(0);
			XSSFCell out_cell = out_row.createCell(0);
			out_cell.setCellValue("Families");
			for(int y = 0; y < myFiles.size()/2; y++)
			{
				out_cell = out_row.createCell(y+1);
				out_cell.setCellValue(GBM + (y+1));
			}
			out_cell = out_row.createCell(myFiles.size()/2 +2);
			out_cell.setCellValue("Total");
			
			int x = 2; //start at row 3 (gives space after header)
			for(int y = 0; y < all_maps.length; y++)
			{
				//write current family name
				out_row = out_sheet.createRow(x);
				x++;
				out_cell = out_row.createCell(0);
				out_cell.setCellValue(family_names[y]);
				out_row = out_sheet.createRow(x);
				x++;
				
				//Iterator to print each name in each family
				Iterator<String> keySetIterator = all_sorted_maps[y].keySet().iterator();
				//Print out the name of each family member
				while(keySetIterator.hasNext())
				{	
					//get next name
					String key = keySetIterator.next();
					
					//write name
					out_row = out_sheet.createRow(x);
					x++;
					out_cell = out_row.createCell(0);
					out_cell.setCellValue(key);
					
					int p_index = 0;
					
					//write out each GBM they attended
					for(int i = 0; i < myFiles.size()/2; i++)
					{
						
						out_cell = out_row.createCell(i+1);
						out_cell.setCellValue(attendance_array[y][all_sorted_maps[y].get(key)][p_index] +
											attendance_array[y][all_sorted_maps[y].get(key)][p_index+1]);
						p_index += 2;
					}
					
					out_cell = out_row.createCell(myFiles.size()/2 +2);
					out_cell.setCellValue(attendance_totals[y][all_sorted_maps[y].get(key)][0]);
					
					
				}
				out_row = out_sheet.createRow(x);
				x++;
			}
			
			//resize first column for longer names
			out_sheet.autoSizeColumn(0);
			//Write the sheet/book to an excel file
			out_book.write(new FileOutputStream("FullAttendance.xlsx"));
			
			//Close everything
			out_book.close();
			
			
	}
}
