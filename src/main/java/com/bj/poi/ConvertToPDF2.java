package com.bj.poi;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;

//needed jars: xdocreport-2.0.1.jar, 
//odfdom-java-0.8.7.jar,
//itext-2.1.7.jar  
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.converter.XDocConverterException;
import fr.opensagres.xdocreport.converter.IConverter;
import fr.opensagres.xdocreport.converter.ConverterRegistry;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.core.document.DocumentKind;

//needed jars: apache poi and it's dependencies
//and additionally: ooxml-schemas-1.3.jar 
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

// ref : https://stackoverflow.com/questions/51330192/trying-to-make-simple-pdf-document-with-apache-poi/51337157#51337157
public class ConvertToPDF2 {
    public void execute(String docPath, String pdfPath) throws InvalidFormatException, IOException, XDocConverterException {
    	  XWPFDocument document = new XWPFDocument();

    	  // there must be a styles document, even if it is empty
    	  XWPFStyles styles = document.createStyles();

    	  // there must be section properties for the page having at least the page size set
    	  CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
//    	  CTPageSz pageSz = sectPr.addNewPgSz();
//    	  pageSz.setW(BigInteger.valueOf(12240)); //12240 Twips = 12240/20 = 612 pt = 612/72 = 8.5"
//    	  pageSz.setH(BigInteger.valueOf(15840)); //15840 Twips = 15840/20 = 792 pt = 792/72 = 11"

    	  // filling the body
    	  XWPFParagraph paragraph = document.createParagraph();

    	  //create table
    	  XWPFTable table = document.createTable();

    	  //create first row
    	  XWPFTableRow tableRowOne = table.getRow(0);
    	  tableRowOne.getCell(0).setText("col one, row one");
    	  tableRowOne.addNewTableCell().setText("col two, row one");
    	  tableRowOne.addNewTableCell().setText("col three, row one");

    	  //create CTTblGrid for this table with widths of the 3 columns. 
    	  //necessary for Libreoffice/Openoffice and PdfConverter to accept the column widths.
    	  //values are in unit twentieths of a point (1/1440 of an inch)
    	  //first column = 2 inches width
    	  table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(2*1440));
    	  //other columns (2 in this case) also each 2 inches width
    	  for (int col = 1 ; col < 3; col++) {
    	   table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2*1440));
    	  }

    	  //create second row
    	  XWPFTableRow tableRowTwo = table.createRow();
    	  tableRowTwo.getCell(0).setText("col one, row two");
    	  tableRowTwo.getCell(1).setText("col two, row two");
    	  tableRowTwo.getCell(2).setText("col three, row two");

    	  //create third row
    	  XWPFTableRow tableRowThree = table.createRow();
    	  tableRowThree.getCell(0).setText("col one, row three");
    	  tableRowThree.getCell(1).setText("col two, row three");
    	  tableRowThree.getCell(2).setText("col three, row three");

    	  paragraph = document.createParagraph();

    	  //trying picture
//    	  XWPFRun run = paragraph.createRun();
//    	  run.setText("The picture in line: ");
//    	  InputStream in = new FileInputStream("samplePict.jpeg");
//    	  run.addPicture(in, Document.PICTURE_TYPE_JPEG, "samplePict.jpeg", Units.toEMU(100), Units.toEMU(30));
//    	  in.close();  
//    	  run.setText(" text after the picture.");

    	  paragraph = document.createParagraph();

    	  //document must be written so underlaaying objects will be committed
    	  ByteArrayOutputStream out = new ByteArrayOutputStream();
    	  document.write(out);
    	  document.close();

    	  // 1) Create options DOCX 2 PDF to select well converter form the registry
    	  Options options = Options.getFrom(DocumentKind.DOCX).to(ConverterTypeTo.PDF);

    	  // 2) Get the converter from the registry
    	  IConverter converter = ConverterRegistry.getRegistry().getConverter(options);

    	  // 3) Convert DOCX 2 PDF
    	  InputStream docxin = new FileInputStream(new File(docPath));
    	  OutputStream pdfout = new FileOutputStream(new File(pdfPath));
    	  converter.convert(docxin, pdfout, options);

    	  docxin.close();       
    	  pdfout.close();
    }
}
