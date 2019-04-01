package com.bj.doc2pdf.convert;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;


public class DocToPDFConverter extends Converter {
	private static final Logger logger = LoggerFactory.getLogger(DocToPDFConverter.class);
	
	public DocToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() {
		// ref : https://stackoverflow.com/questions/6201736/javausing-apache-poi-how-to-convert-ms-word-file-to-pdf
        POIFSFileSystem fs = null;  
        Document document = new Document();		
		try {
			loading();

			InputStream iStream = inStream;
			fs = new POIFSFileSystem(iStream);
			HWPFDocument doc = new HWPFDocument(fs);  
            WordExtractor we = new WordExtractor(doc);  
			
            PdfWriter writer = PdfWriter.getInstance(document, outStream);  

            Range range = doc.getRange();
            document.open();  
            BaseFont bf = BaseFont.createFont(FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            writer.setPageEmpty(true);  
            document.newPage();  
            writer.setPageEmpty(true);  

            String[] paragraphs = we.getParagraphText();  
            for (int i = 0; i < paragraphs.length; i++) {  

                org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);
               // CharacterRun run = pr.getCharacterRun(i);
               // run.setBold(true);
               // run.setCapitalized(true);
               // run.setItalic(true);
                paragraphs[i] = paragraphs[i].replaceAll("\\cM?\r?\n", "");  
                logger.debug("Length:" + paragraphs[i].length());  
                logger.debug("Paragraph" + i + ": " + paragraphs[i].toString());  

                // FIXME bsh - 폰트 사이즈 고정하면 안됨.
	            // add the paragraph to the document  
	            document.add(new Paragraph(paragraphs[i], new Font(bf, 12)));  
            }  
            
        	// ref : https://github.com/plutext/docx4j/issues/286
			// using docx4j -> 오류 발생함.
			// org.docx4j.openpackaging.exceptions.Docx4JException: This file seems to be a binary doc/ppt/xls, not an encrypted OLE2 file containing a doc/pptx/xlsx
//			WordprocessingMLPackage wordMLPackage = getMLPackage(iStream);

			processing();
//			Docx4J.toPDF(wordMLPackage, outStream);

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			document.close();
			finished();
		}

	}
	
	protected WordprocessingMLPackage getMLPackage(InputStream iStream) throws Exception{
		PrintStream originalStdout = System.out;
		
		//Disable stdout temporarily as Doc convert produces alot of output
//		System.setOut(new PrintStream(new OutputStream() {
//			public void write(int b) {
//				//DO NOTHING
//			}
//		}));

		WordprocessingMLPackage mlPackage = new WordprocessingMLPackage();
		mlPackage.load(iStream);
		
//		WordprocessingMLPackage mlPackage = Doc.convert(iStream);
		
		System.setOut(originalStdout);
		return mlPackage;
	}
}
