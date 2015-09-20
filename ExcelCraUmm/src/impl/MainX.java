package impl;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import dao.ReadX;
import exceptions.RowNotFoundException;
import service.CreateX;
import util.ConnectorX;
public class MainX {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		while(true){
		int rNo=0;
		ConnectorX cx;
		CreateX cxx= new CreateX();
		ReadX rx;
		BufferedReader in= new BufferedReader(new InputStreamReader(System.in));
		try {
			System.out.println("Enter file path");
			cxx.setRootpath(in.readLine());
			System.out.println("Enter file name");
			cxx.setFilename(in.readLine());
			System.out.println("Enter row number to be read: ");
			rNo=Integer.parseInt(in.readLine());
			rNo--;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		cx=new ConnectorX();
		rx= new ReadX();
		try {
			cx.init(cxx);	
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("File Not Found or File is not a Excel File Guys");
		}
		try {
			//System.out.println(rx.readR(cx,rNo));
			System.out.println(rx.readC(cx, rNo));
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (RowNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		}
	}

}
