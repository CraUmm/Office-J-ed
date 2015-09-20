/*
 * Reads one Row or Column from the Excel Sheet.
 * 
 */
package dao;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import logic.DistinctCX;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import beans.DataStructures;
import exceptions.RowNotFoundException;
import util.*;
public class ReadX {
	
	public String readR(ConnectorX cx, int rNo) throws EncryptedDocumentException, InvalidFormatException, IOException, RowNotFoundException{
		String fileType=cx.getFileName();
		if(fileType.endsWith("xls")){
			readTopRowH(cx,rNo);
			return "An Excel 2007 file. Using HSSF";
		}
		else if(fileType.endsWith("xlsx")){
			readTopRowX(cx,rNo);
			return "An OOXML file. Using XSSF";
		}
		else{
			return "Not a compatible file";
		}	
	}
	
	public String readC(ConnectorX cx, int cNo) throws EncryptedDocumentException, InvalidFormatException, IOException, RowNotFoundException {
		// TODO Auto-generated method stub
		String fileType=cx.getFileName();
		if(fileType.endsWith("xls")){
			readFirstColumnH(cx, cNo);
			return "An Excel 2007 file. Using HSSF";
		}
		else if(fileType.endsWith("xlsx")){
			readFirstColumnX(cx, cNo);
			return "An OOXML file. Using XSSF";
		}
		else{
			return "Not a compatible file";
		}	
	}
	
	public void readTopRowH(ConnectorX cx,int rNo) throws IOException, EncryptedDocumentException, InvalidFormatException, RowNotFoundException{
		
		HSSFWorkbook wbk= (HSSFWorkbook) WorkbookFactory.create(cx.getFile());
		HSSFSheet xwbksh= wbk.getSheet("Sheet1");
		List<HSSFCell> rowOne =  new DataStructures().getLhCell();
		HSSFRow xRow = xwbksh.getRow(rNo);
		if(xwbksh.getLastRowNum()<rNo){
			throw new RowNotFoundException("Row doesnt exist");
		}
		Iterator<Cell> cells=xRow.iterator();
		System.out.println("Reading the Row Number:"+(xRow.getRowNum()+1));
		String rOne="";
		int i=0;
		//rowOne.add((HSSFCell) cells.getClass());
		while(cells.hasNext()){
			HSSFCell hc= (HSSFCell) cells.next();
			rowOne.add(hc);
			if(hc.getCellType()==Cell.CELL_TYPE_STRING){
				rOne+=rowOne.get(i).getStringCellValue().concat(Constants.fieldGap);
			}
			else if(hc.getCellType()==Cell.CELL_TYPE_NUMERIC){
				rOne+=Double.valueOf(rowOne.get(i).getNumericCellValue()).toString().concat(Constants.fieldGap);
			}
			else if(hc.getCellType()==Cell.CELL_TYPE_FORMULA){
				CellValue cVal= new HSSFFormulaEvaluator(wbk).evaluate(hc);
				rOne+=cVal.formatAsString().concat(Constants.fieldGap);
				//rOne+=Double.valueOf(rowOne.get(i).getCellFormula()).toString().concat(Constants.fieldGap);
			}
			i++;
		}
		System.out.println(rOne);
		cx.getFile().close();
	}
	
	public void readTopRowX(ConnectorX cx,int rNo) throws IOException, EncryptedDocumentException, InvalidFormatException, RowNotFoundException{
		
		XSSFWorkbook wbk= (XSSFWorkbook) WorkbookFactory.create(cx.getFile());
		XSSFSheet xwbksh= wbk.getSheet("Sheet1");
		List<XSSFCell> rowOne =  new DataStructures().getLxCell();
		XSSFRow xRow = xwbksh.getRow(rNo);
		if(xwbksh.getLastRowNum()<rNo){
			throw new RowNotFoundException("Row doesnt exist");
		}
		Iterator<Cell> cells=xRow.iterator();
		System.out.println("Reading the Row Number:"+(xRow.getRowNum()+1));
		String rOne="";
		int i=0;
		//rowOne.add((HSSFCell) cells.getClass());
		while(cells.hasNext()){
			XSSFCell hc= (XSSFCell) cells.next();
			rowOne.add(hc);
			if(hc.getCellType()==Cell.CELL_TYPE_STRING){
				rOne+=rowOne.get(i).getStringCellValue().concat(Constants.fieldGap);
			}
			else if(hc.getCellType()==Cell.CELL_TYPE_NUMERIC){
				rOne+=Double.valueOf(rowOne.get(i).getNumericCellValue()).toString().concat(Constants.fieldGap);
			}
			else if(hc.getCellType()==Cell.CELL_TYPE_FORMULA){
				rOne+=Double.valueOf(rowOne.get(i).getCellFormula()).toString().concat(Constants.fieldGap);
			}
			i++;
		}
		System.out.println(rOne);
		cx.getFile().close();
	}
	
	public void readFirstColumnX(ConnectorX cx,int cNo) throws IOException, EncryptedDocumentException, InvalidFormatException, RowNotFoundException{
		
		int i=0;
		String cOne="";
		XSSFWorkbook wbk = (XSSFWorkbook) WorkbookFactory.create(cx.getFile());
		XSSFSheet xwbksh = wbk.getSheet("Sheet1");
		List<XSSFCell> colOne= new DataStructures().getLxCell();
		
		Iterator<Row> rows = xwbksh.iterator();
		
		while(rows.hasNext()){

			colOne.add((XSSFCell) ((XSSFRow) rows.next()).getCell(cNo));
			if(colOne.get(i).getCellType()==Cell.CELL_TYPE_STRING){
				cOne+=colOne.get(i).getStringCellValue().concat(Constants.fieldGap);
			}
			else if(colOne.get(i).getCellType()==Cell.CELL_TYPE_NUMERIC){
				cOne+=Double.valueOf(colOne.get(i).getNumericCellValue()).toString().concat(Constants.fieldGap);
			}
			else if(colOne.get(i).getCellType()==Cell.CELL_TYPE_FORMULA){
				CellValue cVal= new XSSFFormulaEvaluator(wbk).evaluate(colOne.get(i));
				cOne+=cVal.formatAsString().concat(Constants.fieldGap);
				//cOne+=colOne.get(i).getCellFormula().concat(Constants.fieldGap);
			}
			i++;
		}
		System.out.println(cOne);
		cx.getFile().close();
	}
	
public void readFirstColumnH(ConnectorX cx,int cNo) throws IOException, EncryptedDocumentException, InvalidFormatException, RowNotFoundException{
		
		int i=0;
		String cOne="";
		HSSFWorkbook wbk = (HSSFWorkbook) WorkbookFactory.create(cx.getFile());
		HSSFSheet xwbksh = wbk.getSheet("Sheet1");
		List<HSSFCell> colOne= new ArrayList<HSSFCell>();
		
		Iterator<Row> rows = xwbksh.iterator();
		
		while(rows.hasNext()){

			colOne.add((HSSFCell) ((HSSFRow) rows.next()).getCell(cNo));
			if(colOne.get(i).getCellType()==Cell.CELL_TYPE_STRING){
				cOne+=colOne.get(i).getStringCellValue().concat(Constants.fieldGap);
			}
			else if(colOne.get(i).getCellType()==Cell.CELL_TYPE_NUMERIC){
				cOne+=Double.valueOf(colOne.get(i).getNumericCellValue()).toString().concat(Constants.fieldGap);
			}
			else if(colOne.get(i).getCellType()==Cell.CELL_TYPE_FORMULA){
				cOne+=colOne.get(i).getCellFormula().concat(Constants.fieldGap);
			}
			i++;
		}
		System.out.println(cOne);
		cx.getFile().close();
		new DistinctCX().inihTree((ArrayList<HSSFCell>) colOne,cNo);
	}
}
