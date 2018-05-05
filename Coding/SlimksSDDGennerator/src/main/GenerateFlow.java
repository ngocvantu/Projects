package main;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ShapeTypes;
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

public class GenerateFlow {
	Properties prop = new Properties();
	InputStream input = null;

	{
		try {
			input = new FileInputStream("config.properties");
			prop.load(input);
			System.out.println(prop.getProperty("fileNameFlow"));
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

	private final String FILE_NAME = prop.getProperty("fileNameFlow");
	private final String STEP_FILE = prop.getProperty("stepFile");
	private final int rowOffsetFirst = 1;
	private final int numberOfRowEachTextBox = 2;
	private final int rowSpace = 1;

	private int colStart = 2;
	private int colEnd = 16;

	FileInputStream excelFile;
	XSSFWorkbook workbook = new XSSFWorkbook(XSSFWorkbookType.XLSX);
	XSSFSheet datatypeSheet;

	public static void main(String[] args) throws IOException {
		GenerateFlow check = new GenerateFlow();
		check.check();
	}

	private void check() throws IOException {
		System.out.println("Running.....");

		excelFile = new FileInputStream(new File(FILE_NAME));
		workbook = new XSSFWorkbook(excelFile);
		datatypeSheet = workbook.getSheetAt(0);
		datatypeSheet = workbook.createSheet();
		// datatypeSheet = workbook.createSheet("xin chao");
		// datatypeSheet.setDefaultColumnWidth(3);
		XSSFDrawing drawing = datatypeSheet.createDrawingPatriarch();
		BufferedReader bf = new BufferedReader(new FileReader(new File(STEP_FILE)));
		
		ArrayList<String> lines = new ArrayList<String>();
		String line = null;
		while ((line = bf.readLine()) != null) {
			lines.add(line);
		}
		
		 for (int i = 0; i < lines.size(); i++) {
			 
			 String text = lines.get(i);
		
			 XSSFTextBox textBox1 = drawing.createTextbox(new XSSFClientAnchor(0,  0, 0, 0, colStart, rowOffsetFirst + i * (numberOfRowEachTextBox + rowSpace), colEnd, rowOffsetFirst + i*rowSpace + (i+1)*numberOfRowEachTextBox));
			 textBox1.setText(i + ". " + text);
			 textBox1.setLineStyle(0);
			 textBox1.setLineStyleColor(0, 0, 0);
			 textBox1.setLineWidth(1.5);
			 textBox1.setFillColor(255, 255, 255);
			 textBox1.setWordWrap(true);
			 textBox1.setTextVerticalOverflow(TextVerticalOverflow.ELLIPSIS);
			 textBox1.setBottomInset(0); // bottom margin (similar padding in HTMLL)
			 textBox1.setTopInset(0);
		 }
		
		 XSSFSimpleShape simpleShape = drawing.createSimpleShape(new
		 XSSFClientAnchor(3, 3, 3, 3, 4, 11, 6, 13));
		 simpleShape.setShapeType(ShapeTypes.DIAMOND);
		 simpleShape.setLineStyle(0);
		 simpleShape.setLineStyleColor(0, 0, 0);
		 simpleShape.setLineWidth(1.5);
		// simpleShape.setFillColor(0, 0, 0);
		//
		 XSSFSimpleShape simpleShape2 = drawing.createSimpleShape(new
		 XSSFClientAnchor(3, 3, 3, 3, 5, 3, 5, 4));
		 simpleShape2.setShapeType(ShapeTypes.LINE);
		 simpleShape2.setLineStyle(0);
		 simpleShape2.setLineStyleColor(0, 0, 0);
		 simpleShape2.setLineWidth(1.5);
		 simpleShape2.setFillColor(0, 0, 0);

//		 for (int i = 0; i < 5; i++) {
//		 workbook.cloneSheet(0, "KSC-S-25_2 ��慣�ｿ�ｽ��ｮｪ��惚��ｪ�ｿｽ�ｿ�ｽi��ｿｽ��私�ｿ�ｽ遯ｶ譎｢�ｿ�ｽj��ｽｽd遯ｶ鄂ｵ(" +
//		 (i+2) + "遯ｶ蜃ｪ�ｽ�｡遯ｶ遖ｿ�ｽ)");
//		
//		 workbook.setSheetOrder("KSC-S-25_2 ��慣�ｿ�ｽ��ｮｪ��惚��ｪ�ｿｽ�ｿ�ｽi��ｿｽ��私�ｿ�ｽ遯ｶ譎｢�ｿ�ｽj��ｽｽd遯ｶ鄂ｵ(" +
//		 (i+2) + "遯ｶ蜃ｪ�ｽ�｡遯ｶ遖ｿ�ｽ)", i+1);
//		 XSSFSheet sheetI = workbook.getSheetAt(i);
//		 Cell cellPersonIncharge = sheetI.getRow(0).getCell(15);
//		 System.out.println(cellPersonIncharge.getStringCellValue());
//		 cellPersonIncharge.setCellValue("(TSDV)HungPN");
//		 Cell cellPakage = sheetI.getRow(1).getCell(8);
//		 cellPakage.setCellValue("jp.co.toshiba_sol.slim.ks.substituteinput");
//		 Cell cellDate = sheetI.getRow(0).getCell(13);
//		 cellDate.setCellValue(new Date());
//		 }
		 
		 bf.close();
 
		FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
		workbook.write(output_file);
		output_file.close();
		System.out.println("number of sheet: " + workbook.getNumberOfSheets());
		System.out.println("Finished!!!");
		Desktop desktop = Desktop.getDesktop();
		desktop.open(new File(FILE_NAME));

	}
 

	private int countFilesInDirectory(String directory) {
		File dir = new File(directory);
		return dir.list().length;
	}

}
