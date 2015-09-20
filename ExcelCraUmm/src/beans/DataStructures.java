package beans;

import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class DataStructures {
	private List<XSSFCell> lxCell;
	private List<HSSFCell> lhCell;
	private List<Row> lRow;
	
	public List<HSSFCell> getLhCell() {
		return lhCell;
	}
	public void setLhCell(List<HSSFCell> lhCell) {
		this.lhCell = lhCell;
	}

	
	public List<XSSFCell> getLxCell() {
		return lxCell;
	}
	public void setLxCell(List<XSSFCell> lxCell) {
		this.lxCell = lxCell;
	}
	public List<Row> getlRow() {
		return lRow;
	}
	public void setlRow(List<Row> lRow) {
		this.lRow = lRow;
	}
}
