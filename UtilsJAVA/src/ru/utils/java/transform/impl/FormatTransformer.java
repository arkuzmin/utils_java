package ru.utils.java.transform.impl;

import org.dom4j.Document;

public class FormatTransformer {
	
	private static final String EXCEL_SHEET_NAME = "Лист1";

	public static byte[] xmlToCSV(Document xml, String delimiter) {
		return (new XmlToCsvTransformer(xml, delimiter)).transform();
	}
	
	public static byte[] xmlToExcel(Document xml, ExcelTypes excelType) {
		return (new XmlToExcelTransformer(xml, excelType, EXCEL_SHEET_NAME)).transform();
	}
	
	public static Document csvToXml(byte[] csvBase64, String delimiter) {
		return (new CsvToXmlTransformer(csvBase64, delimiter)).transform();
	}
	
	public static Document excelToXml(byte[] excelBase64, ExcelTypes excelType) {
		return (new ExcelToXmlTransformer(excelBase64, excelType, EXCEL_SHEET_NAME)).transform();
	}
}
