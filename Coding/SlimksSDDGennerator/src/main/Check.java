package main;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.sl.usermodel.Line;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xslf.usermodel.XSLFDrawing;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xssf.usermodel.TextAutofit;
import org.apache.poi.xssf.usermodel.TextVerticalOverflow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import org.openqa.selenium.WebDriver;

public class Check {
	Properties prop = new Properties();
	InputStream input = null;

	{
		try {
			input = new FileInputStream("config.properties");
			prop.load(input);
			System.out.println(prop.getProperty("fileName"));
			System.out.println(prop.getProperty("pic"));
			System.out.println(prop.getProperty("package"));

//			FileUtils.copyFile(new File(prop.getProperty("fileName")),
//					new File("generated_" + prop.getProperty("fileName")));

		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			if (input != null) {
				try {
					input.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	private final String FILE_NAME = prop.getProperty("fileName");
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
		// datatypeSheet = workbook.createSheet("xin chao");
		// datatypeSheet.setDefaultColumnWidth(3);
		XSSFDrawing drawing = datatypeSheet.createDrawingPatriarch(); 

		int numBerOfClass = countFilesInDirectory(prop.getProperty("package"));
		System.out.println("number of class: ----------------->>>>>>>>>>>>" + numBerOfClass);
		System.out.println("first sheet");
		firstSheetInput(prop.getProperty("package"), workbook.getSheetAt(0));

		File dir = new File(prop.getProperty("package"));
		File[] listFile = dir.listFiles();

		// function description  
		functionDescription(prop.getProperty("package"), listFile);

		FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
		workbook.write(output_file);
		output_file.close();
		System.out.println("number of sheet: " + workbook.getNumberOfSheets());
		System.out.println("Finished!!!");
		Desktop desktop = Desktop.getDesktop();
		desktop.open(new File(FILE_NAME));

	}

	private void methodLogicDiagramAndDetail(ArrayList<String> listFunction) {
		ArrayList<XSSFSheet> clonedSheets = new ArrayList();
		for (int i = 0; i < listFunction.size(); i++) {
			XSSFSheet cloned = workbook.cloneSheet(1 + listFunction.size() + 1,
					prop.getProperty("screenID") + "_3メソッドロジック図(" + (i + 1) + "枚目)");
			XSSFSheet cloned2 = workbook.cloneSheet(1 + listFunction.size() + 2,
					prop.getProperty("screenID") + "_3メソッドロジック説明(" + (i + 1) + "枚目)");
			clonedSheets.add(cloned);
			clonedSheets.add(cloned2);
		}
		
		for (int i = 0; i < clonedSheets.size(); i++) {
//			System.out.println(listFunction.get(i));
			String sheetName = clonedSheets.get(i).getSheetName();
			XSSFSheet clonedSheet = clonedSheets.get(i);
			workbook.setSheetOrder(sheetName, i + 1 + listFunction.size());
			
			clonedSheet.getRow(0).getCell(40).setCellValue(prop.getProperty("pic"));
			clonedSheet.getRow(0).getCell(35).setCellValue(prop.getProperty("date"));
			String functionNameArr[] = listFunction.get(i/2).split("\\.");
			String functionName = functionNameArr[1] + "." + functionNameArr[0] + "()関数";
			clonedSheet.getRow(1).getCell(22).setCellValue(functionName);
			clonedSheet.getRow(1).getCell(14).setCellValue(prop.getProperty("screenID"));
			if (i%2==0) {
				clonedSheet.getRow(1).getCell(5).setCellValue(prop.getProperty("screenID") + "_3メソッドロジック図");
			} else {
				clonedSheet.getRow(1).getCell(5).setCellValue(prop.getProperty("screenID") + "_3メソッドロジック説明");
			}
			
		}
	}

	private void functionDescription(String path, File[] listFile) {
		ArrayList<String> newFunctionList = new ArrayList();
		ArrayList<String> functionList = new ArrayList();
		try {
			for (int i = 0; i < listFile.length; i++) {
				String fileName = listFile[i].getName();
				File file = new File(path + "\\" + fileName);
				FileReader fileReader;
				fileReader = new FileReader(file);
				BufferedReader bufferedReader = new BufferedReader(fileReader);

				String line;
				while ((line = bufferedReader.readLine()) != null) {
					if (line.trim().startsWith("function") || (!listFile[i].getName().endsWith("DTO.java")
							&& !listFile[i].getName().endsWith("DS.java")
							&& !listFile[i].getName().endsWith("AF.java") && (line.trim().startsWith("private")
									|| line.trim().startsWith("public") || line.trim().startsWith("protected"))
							&& line.contains("("))) {
						String lineSplitedArr[] = line.split("\\(");
						String lineSplited[] = lineSplitedArr[0].split(" ");
						String functionName = lineSplited[lineSplited.length - 1];
						String functionNameWithType = null;

						String[] listFileSplited = listFile[i].getName().split("\\\\");
						functionNameWithType = functionName + "." + listFileSplited[listFileSplited.length - 1];
						String checkFunctionNameArr[] = functionNameWithType.split("\\.");
						if (!checkFunctionNameArr[0].equals(checkFunctionNameArr[1]) &&
								checkFunctionNameArr.length < 4) {
							System.out.println("function " + (i + 1) + ": " + functionNameWithType);
							newFunctionList.add(functionNameWithType);
							functionList.add(functionNameWithType);
						}
					}
				}

				fileReader.close();
			}
			functionDescriptionInput(functionList);
			methodLogicDiagramAndDetail(newFunctionList);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	/*
	 * with an array of function, this method will insert to SDD the function
	 * description
	 */
	private void functionDescriptionInput(ArrayList<String> functionList) {

		for (int i = 0; i < functionList.size(); i++) {

			String sheetName = prop.getProperty("screenID") + "_2 メソッド（関数）仕様(" + (i + 1) + "枚目)";
			XSSFSheet clonedSheet = workbook.cloneSheet(1, sheetName);
			workbook.setSheetOrder(sheetName, i + 1);
			clonedSheet.getRow(0).getCell(15).setCellValue(prop.getProperty("pic"));
			clonedSheet.getRow(0).getCell(13).setCellType(CellType.STRING);
			clonedSheet.getRow(0).getCell(13).setCellValue(prop.getProperty("date"));
			String packageString = prop.getProperty("package").split(prop.getProperty("javaSource"))[1];  // JavaSource
			packageString = packageString.replaceAll("\\\\", "\\.");
			packageString = packageString.substring(1, packageString.length());
			// package name
			if (functionList.get(i).endsWith("jsp")) {
				clonedSheet.getRow(1).getCell(8).setCellValue("jsp");
			} else if (functionList.get(i).endsWith("java")) {
				clonedSheet.getRow(1).getCell(8).setCellValue(packageString);
			}
			// class name
			clonedSheet.getRow(2).getCell(5)
					.setCellValue(functionList.get(i).split("\\.")[1] + "." + functionList.get(i).split("\\.")[2]);
			// method name
			clonedSheet.getRow(3).getCell(2).setCellValue(functionList.get(i).split("\\.")[0]);
			System.out.println(packageString);
		}
	}

	/**
	 * input data to first sheet of SDD each class is a row in this sheet insert
	 * column: 1, 2, 4, 5, 6
	 * 
	 * @param directory
	 * @param sheet0
	 */
	private void firstSheetInput(String directory, XSSFSheet sheet0) {

		// person in charge and date
		sheet0.getRow(0).getCell(14).setCellValue(prop.getProperty("pic"));
		sheet0.getRow(0).getCell(12).setCellValue(prop.getProperty("date"));

		File dir = new File(directory);
		File[] classes = dir.listFiles();

		for (int i = 0; i < classes.length; i++) {
			CellStyle style = workbook.createCellStyle();
			style.setWrapText(true);
			/*
			 * style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			 * style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			 * sheet0.getRow(4+i).getCell(1).setCellStyle(style);
			 * sheet0.getRow(4+i).setRowStyle(style);
			 * sheet0.getRow(4+i).getCell(1).setCellStyle(style);
			 */
			sheet0.getRow(4 + i).getCell(0).setCellValue(i + 1);
			String classname = classes[i].getName();
			String[] fileID = classname.split("\\.");
			sheet0.getMergedRegions();
			sheet0.getRow(4 + i).getCell(1).setCellValue(fileID[1] + "_" + fileID[0]);
			sheet0.getRow(4 + i).getCell(4).setCellValue(classname);
			String fileType = "";
			if (classname.endsWith("java")) {
				fileType = "JAVA";
			} else if (classname.endsWith("jsp")) {
				fileType = "JSP";
			}
			sheet0.getRow(4 + i).getCell(5).setCellValue(fileType);
			sheet0.getRow(4 + i).getCell(6).setCellValue("実装ファイル");
		} // end for loop

		for (int i = classes.length; i < 100; i++) {
			XSSFRow row = sheet0.getRow(4 + i);
			if (!(row == null)) {
				for (int j = 0; j < 15; j++) {
					XSSFCell cellToDelete = row.getCell(j);
					cellToDelete.setCellValue("");
				}
			}
		}
	}

	private int countFilesInDirectory(String directory) {
		File dir = new File(directory);
		return dir.list().length;
	}

}
