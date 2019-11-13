package pdfconverter;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.filechooser.FileFilter;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.text.PDFTextStripper;
import org.fit.pdfdom.PDFDomTree;
import org.w3c.dom.Document;

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

public class converter2 {
	//static String[] fields = {"Num", "First Name", "Last Name", "SSN", "DOB", "Phone 1", "Phone 2", "Phone 3", "Street 1", "City 1", "State 1", "Zip 1", "Street 2", "City 2", "State 2", "Zip 2", "Street 3", "City 3", "State 3", "Zip 3"};
	static String[] fields = {"First Name", "Last Name", "Address", "City", "State", "Zip", "AccountNumber", "OriginalBalance", "OriginalCreditor", "SocialSecurityNumber", "PrimaryPhone", "WorkPhone", "EmailAddress", "DateChargedOff", "DateAccountOpened", "Custom1", "Custom2", "Custom3", "Custom4", "Custom5", "Custom6", "Custom7", "Custom8", "Custom9", "Custom10", "Custom11", "Custom12", "Custom13",};
	static String regex = "(?!000|666)[0-8][0-9]{2}-(?!00)[0-9]{2}-(?!0000)[0-9]{4}";
	static String regex2 = "\\d+";;
	static WritableWorkbook workbook;
	static String path = System.getProperty("user.dir") + "/pdf";
	static String output = System.getProperty("user.dir") + "/output.xls";
	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws IOException, WriteException, ParserConfigurationException {
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
		stripper.setStartPage(1); //Start extracting from page 3
        stripper.setEndPage(1); //Extract till page 5
		for (int i = 0; i < dirList.length; i++) {
			File f = new File(dirList[i].getPath());
		    pd = PDDocument.load(f);
		    PDFDomTree parser = new PDFDomTree();
		    Document dom = parser.createDOM(pd);
		    System.out.print(dom.getTextContent());
		}

	    workbook.write();
	    workbook.close();
	    System.out.print("Completed");
	}
	
	private static void ExcelStart(WritableSheet sheet) throws IOException {
		try {
		    for (int j = 0; j < fields.length; j++) {
			    sheet.addCell(new Label(j, 0, fields[j]));
		    }
		} catch (WriteException e) {

		}
	}
	
	private static void AddRow(WritableSheet sheet, String[] text, int counter) {
		try {
			//sheet.addCell(new Label(0, counter, Integer.toString(counter)));
			
			// Name and SSN
			for (int i = 0; i<text.length;i++) {
				if (text[i].contains("SS:")) {
					String[] t = text[i].split("SS:");
					t[0] = t[0].trim();
					t[1] = t[1].trim();
					String[] name = t[0].split(" ");
					name[0] = name[0].replaceAll("\\*", "");
					sheet.addCell(new Label(0, counter, name[0]));
					if (name.length == 3) {
						sheet.addCell(new Label(1, counter, name[2]));
					}
					else {
						sheet.addCell(new Label(1, counter, name[1]));
					}
					
					if (t[1].contains("E:")) {
						t[1] = t[1].split("E:")[0];
					}
					t[1] = t[1].replaceAll("\\*", "");
					sheet.addCell(new Label(9, counter, t[1]));
				}
			}
			
			// DOB
			for (int i = 0; i<text.length;i++) {
				if (text[i].contains("DOB:")) {
					String[] s = text[i].split("DOB:");
					String dob = s[1].trim();
					dob = dob.replaceAll("\\s+", "");
					dob = dob.substring(0, 8);
					sheet.addCell(new Label(4, counter, dob));
				}
			}
			
			// Phone
			int phoneCount = 0;
			String phoneString;
			for (int i = 0; i<text.length;i++) {
				if (text[i].contains("PH:")) {
					if (phoneCount == 0) {
						phoneString = text[i].trim();
						String[] nums = phoneString.split("PH:");
						String number = nums[1].replaceAll("\\.", "-");
						sheet.addCell(new Label(10, counter, number.trim()));
						//phoneString = phoneString.replaceAll(".", "-");
						//phoneString = phoneString.replaceAll("PH:", "");
						
						for (int ps = 0; ps < nums.length-1; ps++) {
							sheet.addCell(new Label(10+ps, counter, nums[ps+1].trim().replaceAll("\\.",  "-")));
						}
						
						phoneCount++;
					}
				}
			}
			
			// Street City State Zip
			int streetCount = 0;
			String[] data;
			for (int i = 0; i<text.length;i++) {
				if (text[i].contains("*")) {
					String[] street = {text[i]};
					String s1;
					if (text[i].contains("-")) {
						street = street[0].split(regex);
						s1 = street[0];
					}
					if (text[i].contains("DOB:")) {
						street = street[0].split("DOB:");
						s1 = street[0];
					}
					if (text[i].contains("YOB:")) {
						street = street[0].split("YOB:");
						s1 = street[0];
					}
					if (text[i].contains("RPTD:")) {
						street = street[0].split("RPTD:");
						s1 = street[0];
					}
					if (text[i].contains("E:")) {
						street = street[0].split("E:");
						s1 = street[0];
					}

					s1 = street[0];
					
					Pattern p = Pattern.compile(regex2);
					Matcher m = p.matcher(text[i+1]);
					String zip = "00000";
					if (m.find()) {
						zip = m.group();
					}
					
					try {
						if (streetCount == 0 && !(i<=2)) {
							s1 = s1.replaceAll("\\*", "");
							sheet.addCell(new Label(2, counter, s1.trim()));
							String place = text[i+1].split(zip)[0];
							String state = place.substring(place.length()-3, place.length());
							String city = place.substring(0, place.length()-3);
							//System.out.println(city);
							sheet.addCell(new Label(3, counter, city.trim()));
							sheet.addCell(new Label(4, counter, state.trim()));
							sheet.addCell(new Label(5, counter, zip.trim()));
							streetCount++;
						}
						else if (streetCount == 1) {
							break;/*
							sheet.addCell(new Label(12, counter, s1.trim()));
							String place = text[i+1].split(zip)[0];
							String state = place.substring(place.length()-3, place.length());
							String city = place.substring(0, place.length()-3);
							//System.out.println(city);
							sheet.addCell(new Label(13, counter, city.trim()));
							sheet.addCell(new Label(14, counter, state.trim()));
							sheet.addCell(new Label(15, counter, zip.trim()));
							streetCount++;*/
						}
						else if (streetCount == 2) {
							break;
							/*
							sheet.addCell(new Label(16, counter, s1.trim()));
							String place = text[i+1].split(zip)[0];
							String state = place.substring(place.length()-3, place.length());
							String city = place.substring(0, place.length()-3);
							//System.out.println(city);
							sheet.addCell(new Label(17, counter, city.trim()));
							sheet.addCell(new Label(18, counter, state.trim()));
							sheet.addCell(new Label(19, counter, zip.trim()));
							streetCount++;
							*/
						}
					}
					catch(Exception e) {
						
					}
				}
				if (streetCount == 3) {
					break;
				} 
			}
			
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
