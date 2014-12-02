package ru.utils.java.transform.impl;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStreamReader;
import java.util.LinkedList;
import java.util.List;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;

import ru.utils.java.transform.FromDocumentTransformer;

public class XmlToExcelTransformer implements FromDocumentTransformer {
	
	private final Document xml;
	private final ExcelTypes excelType;
	private String csvTitle;
	private List<String> csvContent;
	private final String EXCEL_SHEET_NAME;
	
	
	public XmlToExcelTransformer(Document xml, ExcelTypes excelType, final String EXCEL_SHEET_NAME) {
		this.xml = xml;
		this.excelType = excelType;
		this.csvContent = new LinkedList<String>();
		this.EXCEL_SHEET_NAME = EXCEL_SHEET_NAME;
	}
	
	public byte[] transform() {
		
		byte[] result = null;
		
		try {
			extractCsvTitleAndContent();
			result = writeToExcelFile();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return Base64.encodeBase64(result);
	}
	
	private void extractCsvTitleAndContent() throws Exception {
		byte[] csv = (new XmlToCsvTransformer(xml, ";")).transform();
		
		ByteArrayInputStream bais = null;
		BufferedReader br = null;
		
		try { 
			bais = new ByteArrayInputStream(csv);
			br = new BufferedReader(new InputStreamReader(bais));
			csvTitle = br.readLine();
			String contentLine = null;
			while ((contentLine = br.readLine()) != null && !"".equals(contentLine)) {
				csvContent.add(contentLine);
			}
		} finally {
			if (br != null) {
				br.close();
			}	
		}
	}
	
	private byte[] writeToExcelFile() throws Exception {
		byte[] result = null;
		ByteArrayOutputStream baos = null;
	    Workbook workbook = null;
	    
		baos = new ByteArrayOutputStream();
		workbook = createWorkbook();
		Sheet excelSheet = workbook.createSheet(EXCEL_SHEET_NAME);
		writeExcelTitle(excelSheet);
		writeExcelContent(excelSheet);
		workbook.write(baos);
		
		result = baos.toByteArray();
		return result;
	}
	
	private Workbook createWorkbook() {
		Workbook wb = null;
		switch (excelType) {
			case XLS:
				wb = new HSSFWorkbook();
				break;
			case XLSX:
				wb = new XSSFWorkbook();
				break;
			default:
				break;
		}
		return wb;
	}
	
	private void writeExcelTitle(Sheet excelSheet) throws Exception {
		String[] titleStrings = csvTitle.split(";", -1);
		if (titleStrings != null) {
			Row row = excelSheet.createRow(0);
			for (int i = 0; i < titleStrings.length; i++) {
				Cell cell = row.createCell(i);
				cell.setCellValue(titleStrings[i]);
			}
		}
	}
	
	private void writeExcelContent(Sheet excelSheet) throws Exception {
		int rowInd = 1;
		for (String contentString : csvContent) {
			Row row = excelSheet.createRow(rowInd);
			String[] contentStrings = contentString.split(";", -1);

			if (contentStrings != null) {
				for (int i = 0; i < contentStrings.length; i++) {
					Cell cell = row.createCell(i);
					cell.setCellValue(contentStrings[i]);
				}
			rowInd++;
			}
		}
	}
}
