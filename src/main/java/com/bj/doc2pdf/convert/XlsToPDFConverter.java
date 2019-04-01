package com.bj.doc2pdf.convert;

import java.io.InputStream;
import java.io.OutputStream;

public class XlsToPDFConverter extends Converter {

	public XlsToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() {
		// FIXME bsh - 로직 구현 필요함.

	}

}
