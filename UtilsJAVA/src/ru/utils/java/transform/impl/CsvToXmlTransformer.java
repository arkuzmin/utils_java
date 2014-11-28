package ru.utils.java.transform.impl;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

import ru.utils.java.transform.ToDocumentTransformer;

public class CsvToXmlTransformer implements ToDocumentTransformer{
	
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
