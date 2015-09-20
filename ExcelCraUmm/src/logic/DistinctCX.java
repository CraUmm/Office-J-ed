package logic;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;

import util.Constants;

public class DistinctCX {
	private Set<String> tsUID;
	
	public void inixTree(ArrayList<XSSFCell> xCell,int cNo) {
		Iterator<XSSFCell> xVal =xCell.iterator();
		while(xVal.hasNext()){
			if(xVal.next().getCellType()==Cell.CELL_TYPE_STRING){
				tsUID.add(xVal.next().getStringCellValue());
			} else if (xVal.next().getCellType()==Cell.CELL_TYPE_NUMERIC) {
				tsUID.add(Double.valueOf(xVal.next().getNumericCellValue()).toString());
			} else {
				tsUID.add(xVal.next().getCellFormula());
			}
		}
		showSet(cNo);
	}

	public void inihTree(ArrayList<HSSFCell> xCell, int cNo) {
		//Iterator<HSSFCell> xVal =xCell.listIterator();
		//while(xVal.hasNext())
		tsUID= new TreeSet<String>();
		for(int i=0;i<xCell.size();i++){
			if(xCell.get(i).getCellType()==Cell.CELL_TYPE_STRING){
				tsUID.add(xCell.get(i).getStringCellValue());
			} else if (xCell.get(i).getCellType()==Cell.CELL_TYPE_NUMERIC) {
				tsUID.add(Double.valueOf(xCell.get(i).getNumericCellValue()).toString());
			} else {
				tsUID.add(xCell.get(i).getCellFormula());
			}
		}
		showSet(cNo);
	}
	
	public void showSet(int cNo){
		String set="";
		Iterator<String> showSet= tsUID.iterator();
		System.out.println("Distinct elements of the Column Number: "+(cNo+1));
		while(showSet.hasNext()){
			System.out.println(showSet.next()+"\n");
		}
		System.out.println(set);
	}
	
	public Set<String> getTsUID() {
		return tsUID;
	}

	public void setTsUID(TreeSet<String> tsUID) {
		this.tsUID = tsUID;
	}
}
