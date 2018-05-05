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
import org.apache.poi.ss.usermodel.VerticalAlignment;
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
	private int colEnd = 8;

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
			 
			 XSSFSimpleShape simpleShape2 = drawing.createSimpleShape(new
					 XSSFClientAnchor(3, 3, 3, 3, (colStart + colEnd)/2, i * (numberOfRowEachTextBox + rowSpace) , (colStart + colEnd)/2, rowOffsetFirst + i*rowSpace + (i+1)*numberOfRowEachTextBox - numberOfRowEachTextBox));
					 simpleShape2.setShapeType(ShapeTypes.LINE);
					 simpleShape2.setLineStyle(0);
					 simpleShape2.setLineStyleColor(0, 0, 0);
					 simpleShape2.setLineWidth(1.5);
					 simpleShape2.setFillColor(0, 0, 0);
			 
			 String text = lines.get(i); 
			 if (text.toUpperCase().startsWith("IF")) {
				 XSSFSimpleShape simpleShape = drawing.createSimpleShape(new XSSFClientAnchor(0,  0, 0, 0, colStart, rowOffsetFirst + i * (numberOfRowEachTextBox + rowSpace), colEnd, rowOffsetFirst + i*rowSpace + (i+1)*numberOfRowEachTextBox));
				 simpleShape.setShapeType(ShapeTypes.DIAMOND);
				 simpleShape.setLineStyle(0);
				 simpleShape.setLineStyleColor(0, 0, 0);
				 simpleShape.setFillColor(255, 255, 255);
				 simpleShape.setLineWidth(1.5);
				 simpleShape.setVerticalAlignment(VerticalAlignment.CENTER);
				 simpleShape.setText(i + ". " + text);
			} else {
				
				
			 XSSFTextBox textBox1 = drawing.createTextbox(new XSSFClientAnchor(0,  0, 0, 0, colStart, rowOffsetFirst + i * (numberOfRowEachTextBox + rowSpace), colEnd, rowOffsetFirst + i*rowSpace + (i+1)*numberOfRowEachTextBox));
			 textBox1.setText(i + ". " + text);
			 textBox1.setLineStyle(0);
			 textBox1.setLineStyleColor(0, 0, 0);
			 textBox1.setLineWidth(1.5);
			 textBox1.setFillColor(255, 255, 255);
			 textBox1.setWordWrap(true);
			 textBox1.setVerticalAlignment(VerticalAlignment.CENTER);
			 textBox1.setTextVerticalOverflow(TextVerticalOverflow.ELLIPSIS);
			 textBox1.setBottomInset(0); // bottom margin (similar padding in HTMLL)
			 textBox1.setTopInset(0);
			}
		 }
		 
		 

		 
		 bf.close();
 
		FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
		workbook.write(output_file);
		output_file.close();
		System.out.println("number of sheet: " + workbook.getNumberOfSheets());
		System.out.println("Finished!!!");
		Desktop desktop = Desktop.getDesktop();
		desktop.open(new File(FILE_NAME));

	} 

}
