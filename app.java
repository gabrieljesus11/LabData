package app;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Main {
	
	public static void main (String [] args) throws Exception{
		List<Object> entrada = getRowsFromExcel("C:\\Users\\gabri\\Documents\\Biogenar 02-11.xlsx");
		writeExcel("Hoja 1", entrada);
	}
	
	private static List<Object> getRowsFromExcel(String filePath) throws Exception{
		FileInputStream fs = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int i = 1;
		Row row = sheet.getRow(i);		
		List<Object> DATA = new ArrayList();
		DataFormatter objDefaultFormat = new DataFormatter();
		
		
		while(row != null) {
			String[] excelStrings = new String[3];
			Cell cell1 = row.getCell(1);
			Cell cell4 = row.getCell(4);
			Cell cell6 = row.getCell(6);
			excelStrings[0] = objDefaultFormat.formatCellValue(cell1);
			excelStrings[1] = objDefaultFormat.formatCellValue(cell4);
			excelStrings[2] = objDefaultFormat.formatCellValue(cell6);
			DATA.add(excelStrings);
			i++;
			row = sheet.getRow(i);
		}
		workbook.close();
		return DATA;
	}
		
	
	
	private static void writeExcel(String nombreHoja, List<Object> DATA) throws Exception{
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet();
		workbook.setSheetName(0, nombreHoja);
		
		String[] headers = new String[]{
			"APELLIDO Y NOMBRE",
			"DNI",
			"FECHA DE NACIMIENTO"
		};
		
		CellStyle headerStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		headerStyle.setFont(font);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		HSSFRow headerRow = sheet.createRow(0);
		for (int i = 0; i <	headers.length; i++) {
			String header = headers[i];
			HSSFCell cell = headerRow.createCell(i);
			cell.setCellStyle(headerStyle);
			cell.setCellValue(header);
		}
		
		for (int i = 0; i < DATA.size(); i++ ) {
			HSSFRow dataRow = sheet.createRow(i+1);
			
			Object[] d = (Object[]) DATA.get(i);
			String ApellidoYNombre = (String) d[0];
			String DNI = (String) d[1];
			String FechaDeNacimiento = (String) d[2];
			
			dataRow.createCell(0).setCellValue(ApellidoYNombre);
			dataRow.createCell(1).setCellValue(DNI);
			dataRow.createCell(2).setCellValue(FechaDeNacimiento);
		}
		
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd-MM-yyyy");
		LocalDate date = LocalDate.now();
		String dateString = dtf.format(date);
		
		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
		FileOutputStream file = new FileOutputStream("C:\\Users\\gabri\\Tabla De Trabajo " + dateString +".xls");
		workbook.write(file);
		file.close();
	}
}

