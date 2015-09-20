package com.tcs;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.AttributedCharacterIterator.Attribute;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

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
import org.xml.sax.SAXException;
import javax.xml.*;

public class MappingRule {

	public static void main(String[] args) throws ParserConfigurationException,
			IOException, SAXException {
		System.out.println("START");
		String RootPath = "C:\\Users\\559580\\Desktop\\ExportMtt\\mts\\";
		String ExcelFolPath = "C:\\Users\\559580\\Desktop\\ExportMtt\\mts\\Excel\\";
		String FileName = "EremitXmlToIso8583.mss";
		int r = FileName.lastIndexOf(".");
		String FileNameExcel = FileName.substring(0, r);

		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
/**Style**/
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet Sheet2 = wb.createSheet("MappingRules");
		
		HSSFCellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		HSSFRow headerRow = Sheet2.createRow(0);
		HSSFCell headerCell = headerRow.createCell((short) 0);
		headerCell.setCellValue("Map Name:" + FileName);
		headerCell.setCellStyle(style);

		HSSFCellStyle cellStyle1 = wb.createCellStyle();
		HSSFFont boldFont = wb.createFont();
		boldFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		cellStyle1.setFont(boldFont);
		cellStyle1.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		cellStyle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		HSSFCellStyle cellStyle2 = wb.createCellStyle();
		HSSFFont boldFont1 = wb.createFont();
		boldFont1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		cellStyle2.setFont(boldFont1);
		cellStyle2.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
/**Style**/
/**Sheet2 Header**/		
		HSSFRow row = Sheet2.createRow(1);
		HSSFCell cell = row.createCell((short) 0);
		cell.setCellValue("Group/Item Name");
		cell.setCellStyle(style);
		cell = row.createCell((short) 1);
		cell.setCellValue("Rule");
		cell.setCellStyle(style);
		cell = row.createCell((short) 2);
		cell.setCellValue("");
		cell.setCellStyle(style);
		cell = row.createCell((short) 3);
		cell.setCellValue("");
		cell.setCellStyle(style);
		cell = row.createCell((short) 4);
		cell.setCellValue("");
		cell.setCellStyle(style);
/**Sheet2 Header**/
			
		
		Sheet2.setColumnWidth((short) 0, (short) (256 * 25));
		Sheet2.setColumnWidth((short) 1, (short) (256 * 25));
		Sheet2.setColumnWidth((short) 2, (short) (256 * 15));
		Sheet2.setColumnWidth((short) 3, (short) (256 * 15));
		Sheet2.setColumnWidth((short) 4, (short) (256 * 15));

		File file = new File(RootPath + FileName);
		Document document,document1,document2 = null;
		 document = builder.parse(file);

		String	OutputTree = null,InputTree = null ; 
		
		NodeList SchemaList = (NodeList) document
				.getElementsByTagName("Schema");
		for (int k = 0; k < SchemaList.getLength(); k++) {
			if (SchemaList.item(k).getParentNode().getNodeName()
					.equalsIgnoreCase("Output")) {
				OutputTree = SchemaList.item(k).getAttributes().getNamedItem("typetree").getTextContent();

			}
			if (SchemaList.item(k).getParentNode().getNodeName()
					.equalsIgnoreCase("Input")) {
				InputTree = SchemaList.item(k).getAttributes().getNamedItem("typetree").getTextContent();
			}

		}
		r = OutputTree.lastIndexOf(".");
		OutputTree= OutputTree.substring(0, r);
		document1 = builder.parse(new File(RootPath + OutputTree + ".mts"));
		
	
		HSSFSheet Sheet1 = wb.createSheet("MappingRules");		
	/***Added***/	/**Sheet1 Header**/
		HSSFRow row2 = Sheet1.createRow(1);
		HSSFCell cell1 = row2.createCell((short)0);
		cell1.setCellValue("Name");
		cell1.setCellStyle(style);
		cell1 = row.createCell((short)1);
		cell1.setCellValue("Subclass");
		cell1.setCellStyle(style);
		cell1 = row.createCell((short)2);
		cell1.setCellValue("Parent");
		cell1.setCellStyle(style);
		cell1 = row.createCell((short)3);
		cell1.setCellValue("Format");
		cell1.setCellStyle(style);
		cell1 = row.createCell((short)4);
		cell1.setCellValue("");
		cell1.setCellStyle(style);
/***Added***/		/**Sheet1 Header**/	
		
		
		
		
		
		
		
		
//		r = InputTree.lastIndexOf(".");
//		InputTree= InputTree.substring(0, r);
//		document2=builder.parse(new File(RootPath + InputTree + ".mts"));
		
	NodeList ITEM_outputMTS = document1.getElementsByTagName("ITEM");
	
	NodeList GROUP_outputMTS = document1.getElementsByTagName("GROUP");
	NodeList Cat_outputMTS = document1.getElementsByTagName("CATEGORY");
	
	System.out.println(ITEM_outputMTS.getLength());
	System.out.println(GROUP_outputMTS.getLength());
	System.out.println(Cat_outputMTS.getLength());
	
/***Read Rules***/		
		NodeList MapNodeList = (NodeList) document
				.getElementsByTagName("MapRule");
		/** Log **/
		System.out.println(MapNodeList.getLength());
		int universalRowNo = 2;
		HSSFRow row1;
		String Item_Group, Rule = "";
		for (int i = 0; i < MapNodeList.getLength(); i++) {
			row1 = Sheet2.createRow(universalRowNo);
			NodeList MapNodeChild = (NodeList) MapNodeList.item(i);
			for (int j = 0; j < MapNodeChild.getLength(); j++) {

				if (MapNodeChild.item(j).getNodeName()
						.equalsIgnoreCase("objectset")) {
					Item_Group = MapNodeChild
							.item(j)
							.getTextContent()
							.substring(
									0,
									MapNodeChild.item(j).getTextContent()
											.indexOf(":"));
					cell = row1.createCell((short) 0);
					
					
					
					//cell.setCellStyle(cellStyle1);
					cell.setCellValue(Item_Group);

				} else if (MapNodeChild.item(j).getNodeName()
						.equalsIgnoreCase("objectrule")) {
					Rule = MapNodeChild.item(j).getTextContent();
					cell = row1.createCell((short) 1);
					cell.setCellValue(Rule);
				}

			}
			universalRowNo++;
		}
/***Read Rules***/	
		
		
		
		/** Output To Excel **/

		String excelFileName = "Generic.xls";
		FileOutputStream output = new FileOutputStream(new File(ExcelFolPath
				+ FileNameExcel + ".xls"));
		wb.write(output);
		output.flush();
		output.close();

		System.out.println("END");

	}
	
	


}
