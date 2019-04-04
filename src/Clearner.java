import java.util.*;
import java.io.*;
import java.nio.charset.Charset;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Clearner {
	public static void main(String[] args) throws Exception {
		List<List<String>> records = new ArrayList<>();
		String osNewLine = System.getProperty("line.separator");
		String filepath = "C:\\Users\\linsi\\Desktop\\DLDataClearner\\Book1.xlsx";
		
		File excelFile = new File(filepath);
		//InputStream fis = new FileInputStream(excelFile);
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		
	    System.out.println("sheet number of sheet : " + workbook.getNumberOfSheets());
	    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
	    	Sheet sheet = workbook.getSheetAt(i);
        	int row = sheet.getLastRowNum();
            int col = sheet.getRow(0).getLastCellNum();
            
            System.out.println("reading sheet..." + "row: " + row + " col: " + col);
            
            for(int r = 0; r < row; r++) {
                for(int c = 0; c < col; c++) {
                    String text = sheet.getRow(r).getCell(c).getStringCellValue();
                    if (records.size() >= r) {
                		records.add(new ArrayList<>());
                	}
                    if (r == 0) {
                    	text = text.substring(0, text.length() - 2);
                    } else {
                    	text = text.split(",")[0];
                    }
                    records.get(r).add(text);
                }
            }
        }
        
        System.out.println("done reading");
        
        //write to csv
		FileOutputStream outputStream = new FileOutputStream(filepath + ".csv");
		for (int i = 0; i < records.size(); i++) {
			for (int j = 0; j < records.get(i).size(); j++) {
				String newLine = records.get(i).get(j);
				if (j != records.get(i).size() - 1) {
					newLine += ",";
				} else {
					newLine += osNewLine;
				}

				outputStream.write(newLine.getBytes(Charset.forName("UTF-8")));
			}
		}
		outputStream.close();
		
		System.out.println("done writting");
	}
}