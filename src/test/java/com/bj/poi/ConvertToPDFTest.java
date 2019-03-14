package com.bj.poi;

import static org.junit.Assert.*;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import fr.opensagres.xdocreport.converter.XDocConverterException;

public class ConvertToPDFTest {

	@Test
	public void test_doc() {
		ConvertToPDF2 convPdf = new ConvertToPDF2();
        System.out.println("Start");
        try {
			convPdf.execute("/Users/baesunghan/Downloads/doc01.doc", "/Users/baesunghan/Downloads/doc01.pdf");
		} catch (InvalidFormatException | XDocConverterException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		fail("Not yet implemented");
	}
	@Test
	public void test_docx() {
		ConvertToPDF2 convPdf = new ConvertToPDF2();
		System.out.println("Start");
		try {
			convPdf.execute("/Users/baesunghan/Downloads/doc02.docx", "/Users/baesunghan/Downloads/doc02.pdf");
		} catch (InvalidFormatException | XDocConverterException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		fail("Not yet implemented");
	}

}
