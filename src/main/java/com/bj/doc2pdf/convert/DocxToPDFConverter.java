package com.bj.doc2pdf.convert;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.pdf.BaseFont;

import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.converter.ConverterRegistry;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.IConverter;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.converter.XDocConverterException;
import fr.opensagres.xdocreport.core.document.DocumentKind;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;

public class DocxToPDFConverter extends Converter {

	public DocxToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() {
		loading();

		try {
			XWPFDocument document = new XWPFDocument(inStream);
	
			// https://github.com/opensagres/xdocreport.samples/blob/master/samples/fr.opensagres.xdocreport.samples.docx.converters/src/fr/opensagres/xdocreport/samples/docx/converters/pdf/ConvertDocxStructuresToPDF.java
			
			// // PDFViaITextOptions.create().fontEncoding( "windows-1250" );
			Options options = Options.getFrom(DocumentKind.DOCX).to(ConverterTypeTo.PDF);
			
			// ref : https://github.com/opensagres/xdocreport/issues/129
			PdfOptions pdfOptions = PdfOptions.create();
			pdfOptions.fontProvider(new IFontProvider() {

			    @Override
			    public Font getFont(String familyName, String encoding, float size, int style, Color color) {
				    try {
	//			        if (familyName.equalsIgnoreCase("CALIBRI")) {
	//			            BaseFont baseFont =
	//			                    BaseFont.createFont(fontLibrary + "CALIBRI.TTF", encoding, BaseFont.EMBEDDED);
	//			            return new Font(baseFont, size, style, color);
	//
	//			        } else if (familyName.equalsIgnoreCase("CAMBRIA")) {
	//			            BaseFont baseFont =
	//			                    BaseFont.createFont(fontLibrary + "CAMBRIA.TTF", encoding, BaseFont.EMBEDDED);
	//			            return new Font(baseFont, size, style, color);
	//			        } else {
				            BaseFont baseFont =
				                    BaseFont.createFont(FONT, encoding, BaseFont.EMBEDDED);
				            return new Font(baseFont, size, style, color);
//				    	}
				    } catch (Exception e) {
				        throw new RuntimeException(e);
				    }

//				    return FontFactory.getFont(familyName, encoding, size, style, color);
			    }
			});
			options.subOptions(pdfOptions);
			
			processing();
			PdfConverter.getInstance().convert(document, outStream, pdfOptions);
			
//			IConverter converter = ConverterRegistry.getRegistry().getConverter(options);
//			converter.convert(inStream, outStream, options);
			
		} catch (XWPFConverterException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			finished();
		}


	}

}
