package com.bj.tika;

import java.io.IOException;

import org.apache.tika.exception.TikaException;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.Parser;
import org.junit.Test;
import org.xml.sax.SAXException;

import com.bj.tika.extraction.AutoDetectionExample;

public class AutoDetectionExampleTest {

	@Test
	public void extractFromDocFile_워드파일_추출_Test() {
		Parser parser = new AutoDetectParser();
		try {
			AutoDetectionExample.extractFromFile(parser, "/doc01.doc");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TikaException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@Test
	public void extractFromDocxFile_워드파일_추출_Test() {
		Parser parser = new AutoDetectParser();
		try {
			AutoDetectionExample.extractFromFile(parser, "/doc02.docx");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TikaException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@Test
	public void extractFromPptxFile_PPT파일_추출_Test() {
		Parser parser = new AutoDetectParser();
		try {
			AutoDetectionExample.extractFromFile(parser, "/ppt01.pptx");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TikaException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	@Test
	public void extractFromXlsxFile_엑셀파일_추출_Test() {
		Parser parser = new AutoDetectParser();
		try {
			AutoDetectionExample.extractFromFile(parser, "/excel01.xlsx");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TikaException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@Test
	public void extractFromHwpFile_아래한글파일_추출_Test() {
		Parser parser = new AutoDetectParser();
		try {
			AutoDetectionExample.extractFromFile(parser, "/hwp01.hwp");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TikaException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
