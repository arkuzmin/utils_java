package ru.utils.java.transform.impl;

import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;

import ru.utils.java.transform.ToDocumentTransformer;

public class ExcelToXmlTransformer implements ToDocumentTransformer {
	
	private static final String CSV_DELIMITER = ";";
	private final String EXCEL_SHEET_NAME;
	
	private final byte[] excel;
	private final ExcelTypes excelType;
	private String csvTitle;
	private List<String> csvContent;
	private byte[] csv;
	
	public ExcelToXmlTransformer(byte[] excel, ExcelTypes excelType, final String EXCEL_SHEET_NAME) {
		this.excel = excel;
		this.excelType = excelType;
		this.csvContent = new LinkedList<String>();
		this.EXCEL_SHEET_NAME = EXCEL_SHEET_NAME;
	}
	
	public Document transform() {
		Document xml = null;
		try {
			transformExcelToCsv();
			writeCsv();
			xml = (new CsvToXmlTransformer(csv, CSV_DELIMITER)).transform();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return xml;
	}
	
	private void transformExcelToCsv() throws Exception {
		ByteArrayInputStream bais = new ByteArrayInputStream(excel);
		Workbook wb = createWorkbook(bais);
		Sheet sheet = wb.getSheet(EXCEL_SHEET_NAME);
		
		readCsvTitle(sheet);
		readCsvContent(sheet);
	}
	
	private void readCsvTitle(Sheet sheet) {
		Row row = sheet.getRow(0);
		StringBuilder sb = new StringBuilder(CSV_DELIMITER);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i);
			String content = cell.getStringCellValue();
			sb.append(content).append(CSV_DELIMITER);
		}
		sb.deleteCharAt(sb.length() - 1);
		csvTitle = sb.toString();
	}
	
	private void readCsvContent(Sheet sheet) {
		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			StringBuilder sb = new StringBuilder(";");
			for (int j = 0; j < row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				String content = cell.getStringCellValue();
				sb.append(content).append(";");
			}
			sb.deleteCharAt(sb.length() - 1);
			csvContent.add(sb.toString());
		}
	}
	
	private void writeCsv() throws Exception {
		ByteArrayOutputStream baos = null; 
		BufferedWriter bw = null;
		
		try {
			baos = new ByteArrayOutputStream();
			bw = new BufferedWriter(new OutputStreamWriter(baos));
			writeCsvTitle(bw);
			writeCsvContent(bw);
			bw.flush();
			csv = baos.toByteArray();
		} finally {
			bw.close();
		}
	}
	
	private void writeCsvTitle(BufferedWriter bw) throws IOException {
		bw.write(csvTitle);
		bw.newLine();
	}
	
	private void writeCsvContent(BufferedWriter bw) throws IOException {
		for (String s : csvContent) {
			bw.write(s);
			bw.newLine();
		}
	}
	
	private Workbook createWorkbook(InputStream is) throws Exception {
		Workbook wb = null;
		switch (excelType) {
			case XLS:
				wb = new HSSFWorkbook(is);
				break;
			case XLSX:
				wb = new XSSFWorkbook(is);
				break;
			default:
				break;
		}
		return wb;
	}
}
