package com.bj.doc2pdf.convert;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class XlsxToPDFConverter extends Converter {

	public XlsxToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
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
            
			Document document = new Document();

			PdfWriter writer = PdfWriter.getInstance(document, outStream);
			document.open();

			PdfPTable table = new PdfPTable(2);
			PdfPCell table_cell;
			
			// ref : https://coderanch.com/t/537315/java/Apache-POI-HSSFRow-incompatible-Row
			// ref : https://github.com/nakazawaken1/Excel-To-PDF-with-POI-and-PDFBox/blob/master/src/main/java/jp/qpg/ExcelTo.java
			// ref : https://stackoverflow.com/questions/26056485/java-apache-poi-excel-save-as-pdf
			XSSFWorkbook workbook = new XSSFWorkbook(inStream);
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Header header = sheet.getHeader();
				Iterator<Row> iter = sheet.rowIterator();
				while (iter.hasNext()) {
					Row row = iter.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell1 = cellIterator.next();
						switch (cell1.getCellType()) {
							case BOOLEAN:
								System.out.print(cell1.getBooleanCellValue() + "\n");
								// TODO bsh - 로직 추가 필요함.
								break;
							case NUMERIC:
								System.out.print(cell1.getNumericCellValue() + "\n");
								// TODO bsh - 로직 추가 필요함.
								break;
							case STRING:
								System.out.print(cell1.getStringCellValue() + "\n");
								table_cell = new PdfPCell(new Phrase(cell1.getStringCellValue(), font));
								table.addCell(table_cell);
								// TODO bsh - Cell Color, Font Size, 이미지 설정 로직 추가 필
								break;
						}
					}
				}
			}
			document.add(table);
			
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

}
