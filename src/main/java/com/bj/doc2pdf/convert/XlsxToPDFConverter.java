package com.bj.doc2pdf.convert;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
								break;
							case NUMERIC:
								System.out.print(cell1.getNumericCellValue() + "\n");
								break;
							case STRING:
								System.out.print(cell1.getStringCellValue() + "\n");
								break;
						}
					}
				}
			}
			finished();
		} catch (Exception e) {
			// FIXME bsh - Excpetion 처리 보강 필요함.
			e.printStackTrace();
		}
		
		// FIXME bsh - 리턴값 필요함. test code 에서 성공여부 판단 필요함.
	}

}
