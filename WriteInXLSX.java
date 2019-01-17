package com.read_cell;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class WriteInXLSX {

	public static void main(String[] args) throws Exception {
		
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		XSSFSheet spreadsheet = workbook.createSheet(" WriteData ");
		XSSFRow row;
		ArrayList<String> list=new ArrayList<String>();
		list.add("Ravi");
		list.add("Vijay");  
		list.add("Ravi");  
		list.add("Ajay");
		
		Collection<String> noDups = new HashSet<String>(list);
		
		Iterator i=noDups.iterator(); 
		
		int rowid = 0;
		while(i.hasNext()){  
			row = spreadsheet.createRow(rowid++);
			int cellid = 0;
			while(i.hasNext()){
				 Cell cell = row.createCell(cellid++);
		            cell.setCellValue((String)i.next());
			}
		}

	      //Write the workbook in file system
	      FileOutputStream out = new FileOutputStream(new File("C:/Users/v.sudalaimani/Desktop/read.xlsx"));
	      workbook.write(out);
	      out.close();
	      System.out.println("Writesheet.xlsx written successfully");
		
		
	}


}
