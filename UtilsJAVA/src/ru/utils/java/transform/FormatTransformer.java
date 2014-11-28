package ru.utils.java.transform;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

public class FormatTransformer {
	
	private static final String EXCEL_SHEET_NAME = "Лист1";
	
	private static class ExcelToXmlTransformer {
		
		private static final String CSV_DELIMITER = ";";
		
		private final byte[] excel;
		private final ExcelTypes excelType;
		private String csvTitle;
		private List<String> csvContent;
		private byte[] csv;
		
		public ExcelToXmlTransformer(byte[] excel, ExcelTypes excelType) {
			this.excel = excel;
			this.excelType = excelType;
			this.csvContent = new LinkedList<String>();
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
	
	private static class XmlToExcelTransformer {
		
		private final Document xml;
		private final ExcelTypes excelType;
		private String csvTitle;
		private List<String> csvContent;
		
		
		public XmlToExcelTransformer(Document xml, ExcelTypes excelType) {
			this.xml = xml;
			this.excelType = excelType;
			this.csvContent = new LinkedList<String>();
		}
		
		public byte[] transform() {
			
			byte[] result = null;
			
			try {
				extractCsvTitleAndContent();
				result = writeToExcelFile();
			} catch (Exception e) {
				e.printStackTrace();
			}
			
			return result;
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
	
	private static class XmlToCsvTransformer {
		
		private final Document xml;
		private final String delimiter;
		private final List<String> title;
		private final Map<String, Integer> elementIndexes;
		private final List<Map<String, String>> content;
		private String prevElementTitle;
		Map<String, String> prevCsvLine;
		
		public XmlToCsvTransformer (Document xml, String delimiter) {
			this.xml = xml;
			this.delimiter = delimiter;
			this.title = new LinkedList<String>();
			this.elementIndexes = new HashMap<String, Integer>();
			this.content = new LinkedList<Map<String, String>>();
			this.prevCsvLine = null;
		}
		
		public byte[] transform() {
			parseDocumentRecursively(xml.getRootElement());
			return getResultsAsByteArray();
		}
		
		private void parseDocumentRecursively(Element element) {
			@SuppressWarnings("unchecked")
			List<Element> childElements = (List<Element>) element.elements();
			String elementTitle = element.getPath();
			
			if (isIndexElement(childElements)) {
				elementTitle = getIndexTitle(elementTitle);
				updateElementAndHisChildrenIndexes(elementTitle);
			} else {
				Map<String, String> currCsvLine = formCsvLine(element, elementTitle);
				if (prevCsvLine == null || !joinCsvLine(currCsvLine)) {
					content.add(currCsvLine);
					prevCsvLine = currCsvLine;
				}
			}
			
			setElementTitlePosition(elementTitle);
			prevElementTitle = elementTitle;
			visitChildElements(childElements);
		}
		
		private boolean isIndexElement(List<Element> childElements) {
			return childElements.size() > 0;
		}
		
		private String getIndexTitle(String title) {
			return title == null ? null : "_".concat(title);
		}
		
		private void updateElementAndHisChildrenIndexes(String elementTitle) {
			elementIndexes.put(elementTitle, elementIndexes.get(elementTitle) == null ? 1 : elementIndexes.get(elementTitle) + 1);
			updateElementChildrenIndexes(elementTitle);
		}
		
		private void updateElementChildrenIndexes(String elementTitle) {
			for (String childTitle : title) {
				if (childTitle.startsWith(elementTitle) && !childTitle.equals(elementTitle))
					elementIndexes.put(childTitle, null);
			}
		}
		
		private Map<String, String> formCsvLine(Element element, String elementTitle) {
			Map<String, String> csvLine = new HashMap<String, String>();
			csvLine.put(elementTitle, element.getText());
			for (String t : title) {
				if (t.startsWith("_") && elementTitle.startsWith( t.substring(1) )) {
					csvLine.put(t, elementIndexes.get(t) == null ? null : String.valueOf(elementIndexes.get(t)));
				}
			}
			return csvLine;
		}
		
		private boolean joinCsvLine(Map<String, String> currCsvLine) {
			if (prevCsvLine == null) return false;
			boolean canMerge2Lines = true;
			for (String t : title) {
				String prevValue = prevCsvLine.get(t);
				String currValue = currCsvLine.get(t);
				if (!(canMerge2Lines = canMergeValues(t, prevValue, currValue))) {
					break;
				}
			}
			
			if (canMerge2Lines) {
				for (String key : currCsvLine.keySet()) {
					prevCsvLine.put(key, currCsvLine.get(key));
				}
			}
			
			return canMerge2Lines;
		}
		
		private boolean canMergeValues(String title, String prevValue, String currValue) {
			boolean result = false;
			if (title.startsWith("_")) {
				result = (String.valueOf(prevValue).equals(String.valueOf(currValue)));
			} else {
				result = (String.valueOf(prevValue).equals(String.valueOf(currValue))) ||
						 prevValue == null || currValue == null;
			}
			return result;
		}
		
		
		private void setElementTitlePosition(String elementTitle) {
			if (title.indexOf(elementTitle) == -1) {
				int prevElementIndex = title.indexOf(prevElementTitle);
				int indexToInsert = prevElementIndex == -1 ? 0 : prevElementIndex + 1;
				title.add(indexToInsert , elementTitle);
			}
		}
		
		private void visitChildElements(List<Element> childElements) {
			for (Element child : childElements) {
				parseDocumentRecursively(child);
			}
		}
		
		private String formTitle() {
			StringBuilder titleSb = new StringBuilder("");
			for (String t : title) {
				titleSb.append(t).append(delimiter);
			}
			titleSb.deleteCharAt(titleSb.length() - 1);
			return titleSb.toString();
		}
		
		private List<String> formContent() {
			List<String> contentStrings = new LinkedList<String>();
			
			for (Map<String, String> c : content) {
				StringBuilder contentSb = new StringBuilder("");
				for (String t : title) {
					contentSb.append(c.get(t) == null ? "" : c.get(t)).append(delimiter);
				}
				contentSb.deleteCharAt(contentSb.length() - 1);
				contentStrings.add(contentSb.toString());
			}
			
			return contentStrings;
		}
	
		private byte[] getResultsAsByteArray() {
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(baos));
			byte[] result = null;
			try {
				writeTitle(bw);
				writeContent(bw);
				baos.flush();
				bw.flush();
				result = baos.toByteArray();
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				closeStreams(bw, baos);
			}
			
			return result;
		}
		
		private void writeTitle(BufferedWriter bw) throws IOException {
			bw.write(formTitle());
			bw.newLine();
		}
		
		private void writeContent(BufferedWriter bw) throws IOException {
			for (String c : formContent()) {
				bw.write(c);
				bw.newLine();
			}
		}
		
		private void closeStreams(Writer w, OutputStream os) {
			try {
				if (w != null) {
					w.close();
				}
				if (os != null) {
					os.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	private static class CsvToXmlTransformer {
		
		private final byte[] csv;
		private final Map<String, Integer> elementIndexes;
		private Element currentElement;
		private final Document xml;
		private List<String> csvContent;
		private String csvTitle;
		private final String delimiter;
		
		public CsvToXmlTransformer(byte[] csv, String delimiter) {
			this.csv = csv;
			this.elementIndexes = new HashMap<String, Integer>();
			this.xml = DocumentHelper.createDocument();
			this.currentElement = null;
			this.delimiter = delimiter;
			this.csvContent = new LinkedList<String>();
		}
		
		public Document transform() {
			try {
				extractCsvTitleAndContent();
				List<String> titleValues = getCsvLineValues(csvTitle, delimiter);
				processCsvLines(titleValues);
			} catch (Exception e) {
				e.printStackTrace();
			}
			return this.xml;
		}
		
		private void extractCsvTitleAndContent() throws Exception {
			ByteArrayInputStream bais = new ByteArrayInputStream(csv);
			BufferedReader br = new BufferedReader(new InputStreamReader(bais));
			
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
		
		private List<String> getCsvLineValues(String csvLine, String delimiter) {
			String[] values = csvLine.split(delimiter, -1);
			List<String> result = new LinkedList<String>(Arrays.asList(values));
			return result;
		}
		
		private void processCsvLines(List<String> titleValues) {
			for (String csvLine : csvContent) {
				 List<String> lineValues = getCsvLineValues(csvLine, delimiter);
				 processCsvLine(titleValues, lineValues);
			 }
		}
		
		private void processCsvLine(List<String> titleValues, List<String> lineValues) {
	
			for (int i = 0; i < titleValues.size(); i++) {
				String titleValue = titleValues.get(i);
				String contentValue = lineValues.get(i);
				 
				if (!"".equals(contentValue)) {
					if (isIndexElement(titleValue)) {
						Integer currentIndex = elementIndexes.get(titleValue);
						
						if (currentIndex == null) {
							elementIndexes.put(titleValue, Integer.parseInt(contentValue));	
							currentElement = (currentElement == null) ? xml.addElement(elementTitleWithoutIndexPrefix(titleValue)) 
									 							 	   : currentElement.addElement(elementTitleWithoutIndexPrefix(titleValue));
						} else if (!(currentIndex == Integer.parseInt(contentValue))) {
							updateElementAndHisChildrenIndexes(titleValues, titleValue, contentValue);
							currentElement = (xml.selectSingleNode("//" + elementTitleWithoutIndexPrefix(titleValue) + "[last()]")).getParent().addElement(elementTitleWithoutIndexPrefix(titleValue));
						}
					} else {
						Element el = currentElement.addElement(titleValue);
						el.setText(contentValue);
					}
					
				} else {
					Integer currentIndex = elementIndexes.get(titleValue);
					if (currentIndex != null) {
						currentElement = (xml.selectSingleNode("//" + elementTitleWithoutIndexPrefix(titleValue) + "[last()]")).getParent();
					}
				}
			}
		}
		
		private boolean isIndexElement(String titleValue) {
			return titleValue != null && titleValue.startsWith("_");
		}
		
		private String elementTitleWithoutIndexPrefix(String titleValue) {
			return titleValue.substring(1);
		}
		
		private void updateElementAndHisChildrenIndexes(List<String> titleValues, String titleValue, String contentValue) {
			elementIndexes.put(titleValue, Integer.parseInt(contentValue));	
			int elementPosition = titleValues.indexOf(titleValue);
			for (int j = elementPosition + 1; j < titleValues.size(); j++) {
				elementIndexes.put(titleValues.get(j), null);
			}
		}
		
	}
	
	public static byte[] xmlToCSV(Document xml, String delimiter) {
		return (new XmlToCsvTransformer(xml, delimiter)).transform();
	}
	
	public static byte[] xmlToExcel(Document xml, ExcelTypes excelType) {
		return (new XmlToExcelTransformer(xml, excelType)).transform();
	}
	
	public static Document csvToXml(byte[] csv, String delimiter) {
		return (new CsvToXmlTransformer(csv, delimiter)).transform();
	}
	
	public static Document excelToXml(byte[] excel, ExcelTypes excelType) {
		return (new ExcelToXmlTransformer(excel, excelType)).transform();
	}
	
	public enum ExcelTypes {
		XLS,
		XLSX
	}
}
