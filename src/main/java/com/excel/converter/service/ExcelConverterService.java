package com.excel.converter.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.excel.converter.dto.JsonToExcelDto;
/**
 * 
 * @author Vidya
 *
 */
@Service
public class ExcelConverterService {

	public void jsonToExcelConverter(JsonToExcelDto jsonToExcelDto) throws IOException {
		Workbook workbook = new XSSFWorkbook();

		Sheet sheet = workbook.createSheet("JsonToExcel");
		sheet.setColumnWidth(0, 6000);
		sheet.setColumnWidth(1, 4000);

		int rowNo=0;
		Row header = sheet.createRow(rowNo++);
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		XSSFFont font = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Times New Roman");
		font.setFontHeightInPoints((short) 16);
		font.setBold(true);
		headerStyle.setFont(font);
		int cellNo=0;
		Set<String> keySet= jsonToExcelDto.getHeader().keySet();
		for (String value : jsonToExcelDto.getHeader().values()) {
			Cell headerCell = header.createCell(cellNo++);
			headerCell.setCellValue(value);
			headerCell.setCellStyle(headerStyle);

			
		}
		
		CellStyle style = workbook.createCellStyle();
		style.setWrapText(true);
		
		for (Map<String, String> map : jsonToExcelDto.getValues()) {
			Row row = sheet.createRow(rowNo++);
			cellNo=0;
			for (String key : map.keySet()) {
					Cell cell = row.createCell(new ArrayList<>(keySet).indexOf(key));
					cell.setCellValue(map.get(key));
					cell.setCellStyle(style);
			}

		}
		
		File currDir = new File(".");
		String path = currDir.getAbsolutePath();
		String fileLocation = path.substring(0, path.length() - 1) + "JsonToExcelExample.xlsx";

		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		workbook.write(outputStream);
		workbook.close();

	}

}
