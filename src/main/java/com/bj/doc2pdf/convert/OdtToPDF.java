package com.bj.doc2pdf.convert;

import java.io.InputStream;
import java.io.OutputStream;

import org.odftoolkit.odfdom.doc.OdfTextDocument;

import fr.opensagres.odfdom.converter.pdf.PdfConverter;
import fr.opensagres.odfdom.converter.pdf.PdfOptions;

public class OdtToPDF extends Converter {

	public OdtToPDF(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
		// TODO Auto-generated constructor stub
	}

	@Override
	public void convert() {
		try {
			loading();       


			OdfTextDocument document = OdfTextDocument.loadDocument(inStream);

			PdfOptions options = PdfOptions.create();

			processing();
			PdfConverter.getInstance().convert(document, outStream, options);

			finished();
			
		} catch(Exception e) {
			e.printStackTrace();
		}


	}

}
