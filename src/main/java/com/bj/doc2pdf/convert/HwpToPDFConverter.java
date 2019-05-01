package com.bj.doc2pdf.convert;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.Iterator;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;

import fr.opensagres.xdocreport.converter.ConverterRegistry;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.IConverter;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.core.document.DocumentKind;
import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.writer.HWPWriter;

public class HwpToPDFConverter extends Converter {
		
	public HwpToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() {
		// ref : https://luji.tistory.com/18
		// ref : https://github.com/neolord0/hwplib
		try {
			loading();
			
			processing();

            BaseFont bf = BaseFont.createFont(FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font font = new Font(bf, 12);
            Font subFont = new Font(bf, 10);
            
//			Document document = new Document();

//			PdfWriter writer = PdfWriter.getInstance(document, outStream);
//			document.open();
            
			HWPFile hwpFile = HWPReader.fromInputStream(inStream); 
			
			if (hwpFile != null) {
			
			    // 첫번째 구역/문단에 문자열 추가하고
				Section s = hwpFile.getBodyText().getSectionList().get(0);
				Paragraph firstParagraph = s.getParagraph(0);
				firstParagraph.getText().addString("이것은 추가된 문자열입니다.");

				// 다른 이름으로 저장
				HWPWriter.getOutputStream(hwpFile, outStream);
			}
			
			// Seems like I must close document if not output stream is not complete
//			document.close();

			// Not sure what repercussions are there for closing a writer but just do it.
//			writer.close();
			// // PDFViaITextOptions.create().fontEncoding( "windows-1250" );
			Options options = Options.getFrom(DocumentKind.DOCX).to(ConverterTypeTo.PDF);
			
			IConverter converter = ConverterRegistry.getRegistry().getConverter(options);
			converter.convert(inStream, outStream, options);
			
		} catch (Exception e) {
			// FIXME bsh - Excpetion 처리 보강 필요함.
			e.printStackTrace();
			finished();
		}
		
		// FIXME bsh - 리턴값 필요함. test code 에서 성공여부 판단 필요함.
	}
}
