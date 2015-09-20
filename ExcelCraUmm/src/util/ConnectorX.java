package util;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import service.CreateX;



public class ConnectorX {
	private InputStream filename;
	private String fileName;
	//private XSSFWorkbook workbook;	
	
	public void init(CreateX utilx) throws IOException{
		
		fileName = utilx.getFilename();
		filename =new  FileInputStream(utilx.getRootpath()+"\\"+utilx.getFilename());
		System.out.println("File Found");	

	}
	
	public String getFileName(){
		return fileName;
	}
	
	public InputStream getFile(){
		return filename;
	}
/*	public XSSFSheet giveWorkSheet() throws IOException{
		//XSSFWorkbook ww =new WorkbookFactory().create(filename);
		XSSFWorkbook xf= new XSSFWorkbook(filename);
		XSSFSheet worksheet =xf.getSheet("Sheet1");
		System.out.println(filename.toString());
		return worksheet;
	}*/
}
