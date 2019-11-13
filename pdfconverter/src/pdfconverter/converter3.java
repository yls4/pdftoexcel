package pdfconverter;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.filechooser.FileFilter;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.text.PDFTextStripper;

import jxl.Cell;
//Import the JExcel API
import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.text.PDFTextStripperByArea;

public class converter3 {
	static String[] fields = {"Issue Date", "Amount Due", "Citation Number", "Violation Code Description", "Comment Public", "Is Warning", "License", "Location", "Make", "Officer", "State"};
	static WritableWorkbook workbook;
	static String path = System.getProperty("user.dir") + "/pdf";
	static String output = System.getProperty("user.dir") + "/output.xls";
	@SuppressWarnings("deprecation")
	
	public static void main(String[] args) throws IOException, WriteException {
		workbook = Workbook.createWorkbook(new File(output));
		System.out.println("File created");
		WritableSheet sheet = workbook.createSheet("Page", 0);
		ExcelStart(sheet);
		
		//Scanner user_input = new Scanner( System.in );
		File dir = new File(path);
		//System.out.println(dir.getPath());
		File [] dirList = dir.listFiles(new FilenameFilter() {
		    @Override
		    public boolean accept(File dir, String name) {
		        return name.endsWith(".pdf");
		    }
		});
		
		int counter = 1;
		PDDocument pd;
		PDFTextStripper stripper = new PDFTextStripper();
		PDFTextStripperByArea areaSearch = new PDFTextStripperByArea();
		PDFTextStripperByArea stripper2 = new PDFTextStripperByArea();
		PDFTextStripperByArea stripper3 = new PDFTextStripperByArea();
		//PDRectangle rect = new PDRectangle(0, 0, 100, 100);
		stripper.setStartPage(1); //Start extracting from page 3
        stripper.setEndPage(1); //Extract till page 5
		File f = new File(dirList[0].getPath());
		
	    pd = PDDocument.load(f);
	    //int curHeight = 136;
	    //int rowCount = 37;
	    int curHeight = 116;
	    int rowCount = 39;
	    int rowHeight = 9;
	    int sheetRowCount = 0;
	    int pageStop = 1491;
	    
	    for (int curpage = 800; curpage < pageStop; curpage++) {
	    	if (counter > 800) {
	    		break;
	    	}
	    	PDPage page = pd.getPage(curpage);
	    	
	    	System.out.println("Now parsing page " + curpage);
	    	for (int curRow = 0; curRow < 80; curRow++) {
	    		Rectangle2D.Float cell = new Rectangle2D.Float(0, curHeight, 80, rowHeight);
	    		String name = "cell-1-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		String text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 0, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(80, curHeight, 30, rowHeight);
	    		name = "cell-2-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 1, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(110, curHeight, 40, rowHeight);
	    		name = "cell-3-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 2, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(150, curHeight, 120, rowHeight);
	    		name = "cell-4-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 3, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(270, curHeight, 120, rowHeight);
	    		name = "cell-5-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 4, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(390, curHeight, 40, rowHeight);
	    		name = "cell-6-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 5, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(430, curHeight, 46, rowHeight);
	    		name = "cell-7-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 6, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(476, curHeight, 82, rowHeight);
	    		name = "cell-8-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 7, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(558, curHeight, 65, rowHeight);
	    		name = "cell-9-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 8, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(623, curHeight, 66, rowHeight);
	    		name = "cell-10-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 9, sheetRowCount+1);
	    		
	    		cell = new Rectangle2D.Float(689, curHeight, 100, rowHeight);
	    		name = "cell-11-"+curRow;
	    		areaSearch.addRegion(name, cell);
	    		areaSearch.extractRegions(page);
	    		text = areaSearch.getTextForRegion(name);
	    		areaSearch.removeRegion(name);
	    		AddCell(sheet, text, 10, sheetRowCount+1);
	    		
	    		sheetRowCount++;
	    		curHeight += rowHeight;
	    	}
	    	
			//Rectangle2D.Float issueDate = new Rectangle2D.Float(0, 0, 80, page.getMediaBox().getHeight());
			//stripper2.addRegion("issueDate", issueDate);
			//Rectangle2D.Float amount = new Rectangle2D.Float(80, 0, 30, page.getMediaBox().getHeight());
			//stripper2.addRegion("amount", amount);
			//Rectangle2D.Float citation = new Rectangle2D.Float(110, 0, 40, page.getMediaBox().getHeight());
			//stripper2.addRegion("citation", citation);
			//Rectangle2D.Float violation = new Rectangle2D.Float(150, 0, 120, page.getMediaBox().getHeight());
			//stripper2.addRegion("violation", violation);
	    	//Rectangle2D.Float comment = new Rectangle2D.Float(270, 0, 120, page.getMediaBox().getHeight());
	    	//stripper2.addRegion("comment", comment);
	    	//Rectangle2D.Float warning = new Rectangle2D.Float(390, 0, 40, page.getMediaBox().getHeight());
	    	//stripper2.addRegion("warning", warning);
	    	//Rectangle2D.Float license = new Rectangle2D.Float(430, 0, 46, page.getMediaBox().getHeight());
	    	//stripper2.addRegion("license", license);
	    	//Rectangle2D.Float lot = new Rectangle2D.Float(476, 0, 82, page.getMediaBox().getHeight());
	    	//stripper2.addRegion("lot", lot);
	    	//Rectangle2D.Float make = new Rectangle2D.Float(558, 0, 65, page.getMediaBox().getHeight());
	    	//stripper2.addRegion("make", make);
	    	//Rectangle2D.Float officer = new Rectangle2D.Float(623, 0, 66, page.getMediaBox().getHeight());
	    	//stripper2.addRegion("officer", officer);
			//Rectangle2D.Float state = new Rectangle2D.Float(689, 0, 100, page.getMediaBox().getHeight());
			//stripper2.addRegion("state", state);
			
	    	//stripper2.extractRegions(page);
		    //String text = stripper2.getTextForRegion("license");

	    	//Rectangle2D.Float row = new Rectangle2D.Float(0, 156, 80, 10);
	    	//stripper3.addRegion("row", row);
	    	//stripper3.extractRegions(page);
	    	//String text = stripper3.getTextForRegion("row");
		    //System.out.println(text);
		    counter++;
		    curHeight = 116;
		    rowCount = 39;
	    }
	    //AddRow(sheet, text, counter);
		//counter++;
		pd.close();

	    System.out.println("Data extracted to Excel, parsing through Excel data...");
	    
	    boolean multiline = true;
	    while (multiline) {
	    	multiline = false;
		    for (int row=0; row<sheet.getRows(); row++) {
		    	Cell cell = sheet.getCell(0, row);
		    	if (cell.getContents().length() < 5) {
		    		multiline = true;
		    		WritableCell cell2 = sheet.getWritableCell(4, row-1);
		    		WritableCell cell3 = sheet.getWritableCell(4, row);
		    		String content = cell2.getContents() + cell3.getContents();
		    		content = content.replace("\n", "").replace("\r", "");
		    		Label l = (Label)cell2;
		    		l.setString(content);
		    		sheet.removeRow(row);
		    	}
		    }
	    }
	    
	    System.out.println("Data extraction complete");
	    workbook.write();
	    workbook.close();
	}
	
	private static void ExcelStart(WritableSheet sheet) throws IOException {
		try {
		    for (int j = 0; j < fields.length; j++) {
			    sheet.addCell(new Label(j, 0, fields[j]));
		    }
		} catch (WriteException e) {

		}
	}
	
	private static void AddCell(WritableSheet sheet, String text, int row, int col) throws RowsExceededException, WriteException {
		sheet.addCell(new Label(row, col, text));
	}
}
