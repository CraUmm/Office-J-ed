package util;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class workBook {
	//returns a sheet of the file
	public Sheet workBookReturnCreator(ConnectorX cx) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated constructor stub
		Workbook wk = WorkbookFactory.create(cx.getFile());
		Sheet wks = wk.getSheet("Sheet2");
		return wks;
	}

}
