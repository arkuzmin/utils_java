package ru.utils.java.transform.impl;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.commons.codec.binary.Base64;
import org.dom4j.Document;
import org.dom4j.Element;

import ru.utils.java.transform.FromDocumentTransformer;

public class XmlToCsvTransformer implements FromDocumentTransformer {
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
		return Base64.encodeBase64(getResultsAsByteArray());
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
			if (!(canMerge2Lines = canMergeValues(prevValue, currValue))) {
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
	
	private boolean canMergeValues(String prevValue, String currValue) {
		return 
				(prevValue != null && currValue == null) ||
				(prevValue == null && currValue != null) ||
				(String.valueOf(prevValue).equals(String.valueOf(currValue)));
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
