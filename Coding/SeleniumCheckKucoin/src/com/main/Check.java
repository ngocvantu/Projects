package com.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.sl.usermodel.Line;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xslf.usermodel.XSLFDrawing;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xssf.usermodel.TextAutofit;
import org.apache.poi.xssf.usermodel.TextVerticalOverflow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import org.openqa.selenium.WebDriver;

public class Check {
	private static final String FILE_NAME = "E:\\GITHUB\\chao1.xlsx";
	private final int rowOffsetFirst = 1;
	private final int numberOfRowEachTextBox = 2;
	private final int rowSpace = 1;
	
	private int colStart = 2;
	private int colEnd = 16; 

	FileInputStream excelFile;
	XSSFWorkbook workbook = new XSSFWorkbook(XSSFWorkbookType.XLSX);
	XSSFSheet datatypeSheet;

	public static void main(String[] args) throws IOException {
		Check check = new Check();
		check.check();
	}

	private void check() throws IOException {
		System.out.println("Running....."); 

		excelFile = new FileInputStream(new File(FILE_NAME));
		workbook = new XSSFWorkbook(excelFile);
		datatypeSheet = workbook.getSheetAt(0);
//		datatypeSheet = workbook.createSheet("xin chao");
		datatypeSheet.setDefaultColumnWidth(3);
		XSSFDrawing drawing = datatypeSheet.createDrawingPatriarch();
		for (int i = 0; i < 10; i++) {
			
			XSSFTextBox textBox1 = drawing.createTextbox(new XSSFClientAnchor(0, 0, 0, 0, colStart,
					rowOffsetFirst + i * (numberOfRowEachTextBox + rowSpace), colEnd, rowOffsetFirst + i*rowSpace + (i+1)*numberOfRowEachTextBox));
			textBox1.setText(i + 
					". Create text box and insert to the , createalk asldjkfh alsjkdfh alksjd  asdl;j falsdf alsdjk faslkjdf f");
			textBox1.setLineStyle(0);
			textBox1.setLineStyleColor(0, 0, 0);
			textBox1.setLineWidth(1.5);
			textBox1.setFillColor(255, 255, 255);
			textBox1.setWordWrap(true);
			textBox1.setTextVerticalOverflow(TextVerticalOverflow.ELLIPSIS);
			textBox1.setBottomInset(0); // bottom margin (similar padding in HTMLL)
			textBox1.setTopInset(0);
		} 
		
		XSSFSimpleShape simpleShape = drawing.createSimpleShape(new XSSFClientAnchor(3, 3, 3, 3, 4, 11, 6, 13));
		simpleShape.setShapeType(ShapeTypes.DIAMOND);
		simpleShape.setLineStyle(0);
		simpleShape.setLineStyleColor(0, 0, 0);
		simpleShape.setLineWidth(1.5);
//		simpleShape.setFillColor(0, 0, 0);

		XSSFSimpleShape simpleShape2 = drawing.createSimpleShape(new XSSFClientAnchor(3, 3, 3, 3, 5, 3, 5, 4));
		simpleShape2.setShapeType(ShapeTypes.LINE);
		simpleShape2.setLineStyle(0);
		simpleShape2.setLineStyleColor(0, 0, 0);
		simpleShape2.setLineWidth(1.5);
		simpleShape2.setFillColor(0, 0, 0);
		
		for (int i = 0; i < 5; i++) {
			workbook.cloneSheet(1, "KSC-S-25_2 メソッド（関数）仕様(" + (i+2) + "枚目)");
			XSSFSheet sheetI = workbook.getSheetAt(i + 2);
			Cell cellPersonIncharge = sheetI.getRow(0).getCell(15);
			System.out.println(cellPersonIncharge.getStringCellValue());
			cellPersonIncharge.setCellValue("(TSDV)HungPN");
			Cell cellPakage = sheetI.getRow(1).getCell(8);
			cellPakage.setCellValue("jp.co.toshiba_sol.slim.ks.substituteinput");
			Cell cellDate = sheetI.getRow(0).getCell(13);
			cellDate.setCellValue(new Date());
		}

		FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
		workbook.write(output_file);
		output_file.close();
		System.out.println("Finished!!!");

	}

}
