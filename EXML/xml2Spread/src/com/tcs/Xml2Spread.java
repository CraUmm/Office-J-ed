package com.tcs;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.poi.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import javax.xml.*;

public class Xml2Spread {

	/**
	 * @param args
	 * @throws ParserConfigurationException 
	 * @throws IOException 
	 * @throws SAXException 
	 * @throws XPathExpressionException 
	 */
	/**
	 * @param args
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws IOException
	 * @throws XPathExpressionException
	 */
	public static void main(String[] args) throws ParserConfigurationException, SAXException, IOException, XPathExpressionException {
		System.out.println("Start");
		String RootPath = "C:\\The TCS\\WTX Code\\";
		String ExcelFolPath  = "C:\\The TCS\\WTX Code\\";
		String FileName = "ERemittance.mts";
		int r=FileName.lastIndexOf(".");
		String FileNameExcel = FileName.substring(0, r);
		
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet Sheet1 = wb.createSheet("SheetXml2Spread");
		//HSSFSheet Sheet2 = wb.createSheet("MappingRules");
		Sheet1.setColumnWidth((short)0,(short)(256 * 25));
		Sheet1.setColumnWidth((short)1,(short)(256 * 25));
		Sheet1.setColumnWidth((short)2,(short) (256 * 15));
		Sheet1.setColumnWidth((short)3, (short)(256 * 15));
		Sheet1.setColumnWidth((short)4, (short)(256 * 15));

		/*Sheet2.setColumnWidth((short)0,(short)(256 * 25));
		Sheet2.setColumnWidth((short)1,(short)(256 * 25));
		Sheet2.setColumnWidth((short)2,(short) (256 * 15));
		Sheet2.setColumnWidth((short)3, (short)(256 * 15));
		Sheet2.setColumnWidth((short)4, (short)(256 * 15));*/

		// Read XMl
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
		HSSFCellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		String mapName = "testSheet";

		HSSFRow headerRow = Sheet1.createRow(0);
		HSSFCell headerCell = headerRow.createCell((short)0);
		headerCell.setCellValue("Map Name:" + mapName);
		headerCell.setCellStyle(style);

		HSSFCellStyle cellStyle1 = wb.createCellStyle();
		HSSFFont boldFont = wb.createFont();
		boldFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		cellStyle1.setFont(boldFont);
		cellStyle1.setFillForegroundColor(HSSFColor.ROSE.index);
		cellStyle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		HSSFCellStyle cellStyle2 = wb.createCellStyle();
		HSSFFont boldFont1 = wb.createFont();
		boldFont1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		cellStyle2.setFont(boldFont1);
		cellStyle2.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		
		
		// int universalRowNo=2;
		HSSFRow row = Sheet1.createRow(1);
		HSSFCell cell = row.createCell((short)0);
		cell.setCellValue("Class");
		cell.setCellStyle(style);
		cell = row.createCell((short)1);
		cell.setCellValue("Subclass");
		cell.setCellStyle(style);
		cell = row.createCell((short)2);
		cell.setCellValue("Parent");
		cell.setCellStyle(style);
		cell = row.createCell((short)3);
		cell.setCellValue("Format");
		cell.setCellStyle(style);
		cell = row.createCell((short)4);
		cell.setCellValue("");
		cell.setCellStyle(style);

		int universalRowNo = 2;

		/** Read XML ***/
		//File file = new File("C:\\Users\\Nitin\\Desktop\\JarFies\\mapimportexport.mts");
		File file = new File(RootPath+FileName);
		Document document = builder.parse(file);// Converts file into a document
												// to be parsed

		NodeList nodeList = (NodeList) document.getDocumentElement()
				.getChildNodes();
		
		String a= (String) document.getDocumentElement().getNodeName();
		System.out.println(a);
		NodeList nodes= nodeList.item(1).getChildNodes();
		System.out.println(nodes.getLength());
		
		
		XPathFactory xfactory = XPathFactory.newInstance();
	    XPath xPath = xfactory.newXPath();
	    
	    
	    NodeList list = (NodeList)xPath.evaluate("/TTMAKER/NEWTREE/*[@CategoryOrItemParent = 'Misc XSD']", new InputSource(new FileReader("C:\\The TCS\\WTX Code\\EremitXmlToIso8583.mss")),XPathConstants.NODESET);
	    System.out.println("TRIAL :: " + list.getLength());
		
		int count = nodeList.getLength();
		for (int i = 0; i < count; i++) {
			Node NewTreeNode = nodeList.item(i);

			NodeList childNodes = NewTreeNode.getChildNodes();
			
			for (int j = 0; j < childNodes.getLength(); j++) {

				if (!childNodes.item(j).getNodeName().equalsIgnoreCase("#text")) {
					HSSFRow row1 = Sheet1.createRow(universalRowNo);

					Node currentNode = childNodes.item(j);

					cell = row1.createCell((short)0);
					if (currentNode.getNodeName().equalsIgnoreCase("GROUP")) {

//						cell.setCellValue("GROUP");
//						cell = row1.createCell((short)1);
						cell.setCellStyle(cellStyle1);
						cell.setCellValue(currentNode.getAttributes().getNamedItem("SimpleTypeName").getTextContent());
						
//						cell = row1.createCell((short)2);
//						cell.setCellValue(currentNode.getAttributes().getNamedItem("CategoryOrGroupParent").getTextContent());
//						
						NodeList Innernode = currentNode.getChildNodes();
						for (int k = 0; k < Innernode.getLength(); k++) {
							Node InnerCurnode = Innernode.item(k);
							if (InnerCurnode.getNodeName().equalsIgnoreCase(
									"Sequence")) {
								cell = row1.createCell((short)3);
								
								cell.setCellValue("Sequence");

								NodeList ImExlist = InnerCurnode
										.getChildNodes();
								for (int l = 0; l < ImExlist.getLength(); l++) {
									Node ImNode = ImExlist.item(l);
									cell = row1.createCell((short)4);
								}

								break;
							}
							if (Innernode.item(k).getNodeName()
									.equalsIgnoreCase("Choice")) {
//								cell = row1.createCell((short)3);
//								cell.setCellValue("Choice");

								System.out.println("Group/Choice");
							}
							if (Innernode.item(k).getNodeName()
									.equalsIgnoreCase("Unordered")) {
//								cell = row1.createCell((short)3);
//								cell.setCellValue("Unordered");
								// System.out.println("Group/Unordered");
								break;
							}
						}

					}
					if (currentNode.getNodeName().equalsIgnoreCase("ITEM")) {
//						cell.setCellValue("ITEM");
//						cell = row1.createCell((short)1);
						cell.setCellStyle(cellStyle2);
						cell.setCellValue(currentNode.getAttributes().getNamedItem("SimpleTypeName").getTextContent());
						cell = row1.createCell((short)2);
//						cell.setCellValue(currentNode.getAttributes().getNamedItem("CategoryOrItemParent").getTextContent());
//						 
					}
					if (currentNode.getNodeName().equalsIgnoreCase("CATEGORY")) {

						cell.setCellStyle(cellStyle1);
						cell.setCellValue(currentNode.getAttributes().getNamedItem("SimpleTypeName").getTextContent());
						cell = row1.createCell((short)2);
					}
					universalRowNo++;

				}
			}

		}

		/** out >> excel **/

		String excelFileName = "Generic.xls";
		FileOutputStream output = new FileOutputStream(
				new File(ExcelFolPath+FileNameExcel+".xls"));
		wb.write(output);
		wb.close();
		output.flush();
		output.close();

		System.out.println("END");


	}

	private static short width(short s, int i) {
		// TODO Auto-generated method stub
		return 0;
	}

}
