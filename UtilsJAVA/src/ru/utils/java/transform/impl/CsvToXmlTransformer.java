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
import org.dom4j.Node;
import org.dom4j.tree.DefaultElement;

import ru.utils.java.transform.ToDocumentTransformer;

public class CsvToXmlTransformer implements ToDocumentTransformer{
	
	private final byte[] csv;
	private final Map<String, Integer> elementIndexes;
	private final Document xml;
	private List<String> csvContent;
	private String csvTitle;
	private final String delimiter;
	
	public CsvToXmlTransformer(byte[] csv, String delimiter) {
		this.csv = csv;
		this.elementIndexes = new HashMap<String, Integer>();
		this.xml = DocumentHelper.createDocument();
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
			String elementTitle = titleValues.get(i);
			addElement(elementTitle, titleValues, lineValues);
		}
	}
	
	private void addElement(String elementTitle, List<String> titleValues, List<String> lineValues) {
		String elementValue = lineValues.get(titleValues.indexOf(elementTitle));
		addElementAndCorrectIndexes(elementTitle, elementValue, titleValues, lineValues);
	}
	
	
	private boolean checkAddingNeededAndCorrectIndexes(String elementName, String elementTitle, String elementValue, String xPathToAddElementAt) {
		if (elementValue == null || "".equals(elementValue)) {
			return false;
		}
		
		if (!isIndexElement(elementTitle)) {
			return true;
		} else {
			List nodes = xml.selectNodes(xPathToAddElementAt + "/" + elementName + "[" + elementValue + "]");
			return (nodes == null || nodes.size() == 0) ? true : false;
		}
	}
	
	private void addElementAndCorrectIndexes(String elementTitle, String elementValue, List<String> titleValues, List<String> lineValues) {
		String elementName = getElementNameFromTitle(elementTitle);
		String elementXPath = getElementXPathFromTitle(elementTitle);
		String xPathToAddElementAt = getXPathToAddElementAt(elementXPath, titleValues, lineValues);
		
		if (!checkAddingNeededAndCorrectIndexes(elementName, elementTitle, elementValue, xPathToAddElementAt)) {
			return;
		}
		
		if (("/").equals(xPathToAddElementAt)) {
			xml.addElement(elementName);
		} else {
			DefaultElement parent = (DefaultElement)(xml.selectNodes(xPathToAddElementAt).get(0));
			Element el = parent.addElement(elementName);
			if (!isIndexElement(elementTitle)) {
				el.setText(elementValue);
			}
		}
	}
	
	private String getElementNameFromTitle(String elementTitle) {
		if (elementTitle == null || "".equals(elementTitle)) {
			return null;
		}
		
		String name = elementTitle.substring(elementTitle.lastIndexOf("/") + 1);
		return name;
	}
	
	private String getElementXPathFromTitle(String elementTitle) {
		if (elementTitle == null || "".equals(elementTitle)) {
			return null;
		}
		
		String xPath = elementTitle.substring(0, elementTitle.lastIndexOf("/"));
		return xPath;
	}
	
	private String getXPathToAddElementAt(String titleXPath, List<String> titleValues, List<String> lineValues) {

		if ("_".equals(titleXPath)) {
			return "/";
		}
		
		String inputXPath = titleXPath.startsWith("_") ? titleXPath : "_".concat(titleXPath);
		String xPathToAdd = inputXPath;
		List<String> xPathParts = new LinkedList<String>();
		
		xPathParts.add(inputXPath);
		while (!xPathToAdd.equals("_")) {
			int delimiterIndex = xPathToAdd.lastIndexOf("/");
			String part = inputXPath.substring(0, delimiterIndex);
			xPathToAdd = xPathToAdd.substring(0, delimiterIndex);
			if (!"_".equals(part)) {
				xPathParts.add(part);
			}
		}
		
		for (String part : xPathParts) {
			String xPathIndex = lineValues.get(titleValues.indexOf(part));
			inputXPath = inputXPath.replaceAll(part, part.concat("[").concat(xPathIndex).concat("]"));
		}
		
		return inputXPath.substring(1);
	}
	
	private boolean isIndexElement(String titleValue) {
		return titleValue != null && titleValue.startsWith("_");
	}
}
