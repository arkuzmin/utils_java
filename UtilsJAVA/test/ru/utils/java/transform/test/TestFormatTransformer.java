package ru.utils.java.transform.test;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;

import org.dom4j.Document;
import org.dom4j.io.SAXReader;

import ru.utils.java.transform.impl.FormatTransformer;

public class TestFormatTransformer {
	public static void main(String[] args) throws Exception {
		//String xml = null;
		//xml = "<Root><Element><Field1>234</Field1><F2>55</F2><A><F1>444</F1><F1>22</F1></A><B><F3>899</F3><F1>5888</F1></B></Element><Element><Field1>234</Field1><A><F1>444</F1><F1>22</F1></A><A><F2>55</F2><F1>234</F1></A><A><F3>899</F3><F1>5888</F1></A><B><F3>899</F3><F1>5888</F1></B><B><F3>899</F3><F1>5888</F1></B><A><F3>11899</F3><F1>115888</F1></A><B><F3>899</F3><F1>5888</F1></B></Element></Root>";
		//xml = "<Root><Element><Field1>234</Field1><F2>55</F2><A><F1>444</F1><F1>22</F1></A></Element><Element><Field1>234</Field1><A><F1>444</F1><F1>22</F1></A><A><F2>55</F2><F1>234</F1></A><A><F3>899</F3><F1>5888</F1></A></Element></Root>";
		//xml = "<Stations><Station><AutoID>10</AutoID><Road>[9]</Road><Enterprise>4105</Enterprise><OKATO>[19058]</OKATO><StationCode>17140</StationCode><LongName>КРАСНОЕ-ЭКСПОРТ</LongName><Name>КРАСНОЕ-Э</Name><StationType>[172]</StationType><DateStart>2013-11-27T14:07:51</DateStart><UpdateUser>Первичная загрузка</UpdateUser><Status>Активная</Status><ExternalInfoMap><ExternalInfo><MDMKey>10</MDMKey><Dictionary>Stations</Dictionary><System>NIIAS</System><ExternalKey>3518</ExternalKey></ExternalInfo></ExternalInfoMap></Station><Station><AutoID>11</AutoID><Road>[91]</Road><Enterprise>41051</Enterprise><OKATO>[190581]</OKATO><StationCode>171401</StationCode><LongName>КРАСНОЕ-ЭКСПОРТ1</LongName><Name>КРАСНОЕ-Э1</Name><StationType>[1721]</StationType><DateStart>2014-11-27T14:07:51</DateStart><UpdateUser>Первичная загрузка1</UpdateUser><Status>Активная1</Status><ExternalInfoMap><ExternalInfo><MDMKey>101</MDMKey><Dictionary>Stations</Dictionary><System>NIIAS</System><ExternalKey>35181</ExternalKey></ExternalInfo></ExternalInfoMap></Station></Stations>";
		//xml = "<Staffs><Staff><AutoID>4957</AutoID><TypeOrganization>ОАОЦентральнаяППК</TypeOrganization><Organization>[1]</Organization><Domain>Центральный</Domain><LastName>Федорищев</LastName><FirstName>Алексей</FirstName><Patronymic>Иванович</Patronymic><Position>[41]</Position><MobilePhoneNumber>+79857682177</MobilePhoneNumber><PhoneNumber><Number>+74957682177</Number></PhoneNumber><Email>afedorischev@yandex.ru</Email><DateStart>2013-12-03T09:23:44</DateStart><DateUpdate>2014-02-18T09:57:08</DateUpdate><UpdateUser>Первичнаязагрузка</UpdateUser><Status>Активная</Status></Staff><Staff><AutoID>49572</AutoID><TypeOrganization>ОАОЦентральнаяППК2</TypeOrganization><Organization>[1]2</Organization><Domain>Центральный2</Domain><LastName>Федорищев2</LastName><FirstName>Алексей2</FirstName><Patronymic>Иванович2</Patronymic><Position>[41]2</Position><MobilePhoneNumber>+7985768217722</MobilePhoneNumber><Email>afedorischev@yandex.ru22</Email><DateStart>2013-12-03T09:23:4422</DateStart><DateUpdate>2014-02-18T09:57:0822</DateUpdate><UpdateUser>Первичнаязагрузка2</UpdateUser><Status>Активная2</Status></Staff></Staffs>";
		//Document doc = DocumentHelper.parseText(xml);
		
		File file = new File("K:\\TestXML.xml");
		SAXReader reader = new SAXReader();
    	Document doc = reader.read(file);
		
		byte[] csv = FormatTransformer.xmlToCSV(doc, "\t");
		Document resultDoc = FormatTransformer.csvToXml(csv, "\t");
     //	byte[] csv = FormatTransformer.xmlToExcel(doc, ExcelTypes.XLSX);
     	
     	FileOutputStream fos = new FileOutputStream("K:\\result.csv");
     	BufferedOutputStream bos = new BufferedOutputStream(fos);
     	bos.write(csv);
     	
     	bos.flush();
     	bos.close();
     	
		//FormatTransformer.xmlToExcel(doc, FormatTransformer.ExcelTypes.XLSX);
		//Document result = FormatTransformer.excelToXml(FormatTransformer.xmlToExcel(doc, FormatTransformer.ExcelTypes.XLS), FormatTransformer.ExcelTypes.XLS);
		
		//byte[] b= FormatTransformer.xmlToCSV(doc, ";");
	//	Document d = FormatTransformer.csvToXml(c);
		//System.out.println(c.getTitledContent());
		System.out.println(resultDoc.asXML());
		
	}
}
