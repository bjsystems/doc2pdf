package com.bj.doc2pdf.convert;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.Anchor;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chapter;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Section;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class XlsToPDFConverter extends Converter {

	//Find out number of columns in the excel 
	private static int numberOfColumns;
		
	public XlsToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() {
		// ref : https://stackoverflow.com/questions/26024193/java-lang-classnotfoundexception-org-apache-xmlbeans-xmloptions
		// ref : https://stackoverflow.com/questions/23080945/java-lang-classnotfoundexception-org-apache-xmlbeans-xmlexception
		try {
			loading();
			
			processing();

            BaseFont bf = BaseFont.createFont(FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font font = new Font(bf, 12);
            Font subFont = new Font(bf, 10);
            
			Document document = new Document();

			PdfWriter writer = PdfWriter.getInstance(document, outStream);
			document.open();
			addMetaData(document);
//            addTitlePage(document, font);
            
			Anchor anchor = new Anchor("First Chapter", font);
            anchor.setName("First Chapter");
            
            // Second parameter is the number of the chapter

            Chapter catPart = new Chapter(new Paragraph(anchor), 1);

            Paragraph subPara = new Paragraph("Table", font);
            Section subCatPart = catPart.addSection(subPara);
            addEmptyLine(subPara, 5);
            
			// ref : https://coderanch.com/t/537315/java/Apache-POI-HSSFRow-incompatible-Row
			// ref : https://github.com/nakazawaken1/Excel-To-PDF-with-POI-and-PDFBox/blob/master/src/main/java/jp/qpg/ExcelTo.java
			// ref : https://stackoverflow.com/questions/26056485/java-apache-poi-excel-save-as-pdf
            // ref : https://sites.google.com/site/uvarajjava/home/excel-to-pdf-using-java
			HSSFWorkbook workbook = new HSSFWorkbook(inStream);
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				HSSFSheet sheet = workbook.getSheetAt(i);
				Header header = sheet.getHeader();
				Iterator<Row> iter = sheet.rowIterator();

				int repeatHeaderColumn = 0;  // 0 - repeat
				boolean flag = true;
				PdfPTable table = null;
				
				while (iter.hasNext()) {
					Row row = iter.next();
					int cellNumber = 0;
//					int numberOfColumns = 0;
					if (flag) {
						table = new PdfPTable(row.getLastCellNum());
//						table.setHorizontalAlignment(Element.ALIGN_CENTER);
						table.setWidthPercentage(100);
						table.setSpacingBefore(0f);
						table.setSpacingAfter(0f);
						flag = false;
					}
					
					// For each row, iterate through each columns
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell1 = cellIterator.next();
						switch (cell1.getCellType()) {
							case BOOLEAN:
								System.out.print(cell1.getBooleanCellValue() + "\n");
								// TODO bsh - 로직 추가 필요함.
								break;
							case BLANK:
								table.addCell(" ");
								cellNumber++;
								break;
							case NUMERIC:
								System.out.print(cell1.getNumericCellValue() + "\n");
								cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell1, table, subFont);
								cellNumber++;
								// TODO bsh - 로직 추가 필요함.
								break;
							case STRING:
								System.out.print(cell1.getStringCellValue() + "\n");
								// Header field 반복 표시 처리시 사용함.
								if (repeatHeaderColumn == 0) {
									numberOfColumns = row.getLastCellNum();
									PdfPCell table_cell = new PdfPCell(new Phrase(cell1.getStringCellValue(), subFont));
									table_cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                                    table.addCell(table_cell);
                                    table.setHeaderRows(1);
//				                    if (row.getRowNum() == 0) {
//				                        table_cell.setBackgroundColor(BaseColor.LIGHT_GRAY);
//				                        table_cell.setBorderColor(BaseColor.BLACK);
//				                    }
								} else {
									cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell1, table, subFont);
								}
								cellNumber++;

								// TODO bsh - Cell Color, Font Size, 이미지 설정 로직 추가 필
								break;
						case ERROR:
							break;
						case FORMULA:
							break;
						case _NONE:
							break;
						default:
							break;
						}
					}
					repeatHeaderColumn = 1;
//                    if(numberOfColumns != cellNumber){
//                        for(int k=0; k<(numberOfColumns-cellNumber); k++){
//                            table.addCell(" ");
//                        }
//                    }
				}
				subCatPart.add(table);
				document.add(subCatPart);
			}
			
			// Seems like I must close document if not output stream is not complete
			document.close();

			// Not sure what repercussions are there for closing a writer but just do it.
			writer.close();
			
		} catch (Exception e) {
			// FIXME bsh - Excpetion 처리 보강 필요함.
			e.printStackTrace();
			finished();
		}
		
		// FIXME bsh - 리턴값 필요함. test code 에서 성공여부 판단 필요함.
	}

	private int getLastCellNum(HSSFSheet sheet) {
		int maxCellNum = 0;
		for (int i=0; i<sheet.getLastRowNum(); i++) {
			maxCellNum = (sheet.getRow(i).getLastCellNum() > maxCellNum) ? sheet.getRow(i).getLastCellNum() : maxCellNum;
		}
		
		return maxCellNum;
	}
	
	private static void addMetaData(Document document) {
		document.addTitle("My first PDF");
		document.addSubject("Using iText");
		document.addKeywords("Java, PDF, iText");
		document.addAuthor("Uvaraj");
		document.addCreator("Uvaraj");

	}

	private static void addTitlePage(Document document, Font font)
			throws DocumentException {

		Paragraph preface = new Paragraph();

		// We add one empty line
		addEmptyLine(preface, 1);

		// Lets write a big header
		preface.add(new Paragraph("Title of the document", font));
		addEmptyLine(preface, 1);

		// Will create: Report generated by: _name, _date
		preface.add(new Paragraph("Report generated by: " + "Uvaraj" + ", "
				+ new Date(), font));
		addEmptyLine(preface, 3);

		preface.add(new Paragraph(
				"This document describes something which is very important ",
				font));

		addEmptyLine(preface, 8);

		preface.add(new Paragraph(
				"This document is a preliminary version  ;-).", font));
		document.add(preface);
		// Start a new page
		document.newPage();
	}

	private static void addEmptyLine(Paragraph paragraph, int number) {
		for (int i = 0; i < number; i++) {
			paragraph.add(new Paragraph(" "));
		}
	}
	
	private static int checkEmptyCellAndAddCellContentToPDFTable(int cellNumber, Cell cell, PdfPTable table, Font font) {
		PdfPCell c = null;
		if (cellNumber == cell.getColumnIndex()) {
			if (cell.getCellType() == CellType.NUMERIC) {
				c = new PdfPCell(new Phrase(Double.toString(cell.getNumericCellValue()), font));
			}

			if (cell.getCellType() == CellType.STRING) {
				c = new PdfPCell(new Phrase(cell.getStringCellValue(), font));
			}
			table.addCell(c);
		} else {
			while (cellNumber < cell.getColumnIndex()) {
				table.addCell(" ");
				cellNumber++;
			}

			if (cellNumber == cell.getColumnIndex()) {
				if (cell.getCellType() == CellType.NUMERIC) {
					c = new PdfPCell(new Phrase(Double.toString(cell.getNumericCellValue()), font));
				}

				if (cell.getCellType() == CellType.STRING) {
					c = new PdfPCell(new Phrase(cell.getStringCellValue(), font));
				}
			}
			table.addCell(c);
			cellNumber = cell.getColumnIndex();
		}

		return cellNumber;
	}
}
