package com.bj.convert;

import static org.junit.Assert.*;

import org.junit.Test;

import com.bj.doc2pdf.convert.Converter;
import com.bj.doc2pdf.convert.Parser;
import com.bj.doc2pdf.model.Document;

public class ConverterTest {
	private String resPath = "src/test/resources";
	private String tgtPath = "target";
	
	@Test
	public void test_doc_doc01_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/doc01.doc");
		doc.setOutFilePath(tgtPath + "/doc01.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}

	@Test
	public void test_docx_corrupt_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/corrupt.docx");
		doc.setOutFilePath(tgtPath + "/corrupt.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}
	
	@Test
	public void test_docx_valid_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/valid.docx");
		doc.setOutFilePath(tgtPath + "/valid.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}

	@Test
	public void test_docx_doc02_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/doc02.docx");
		doc.setOutFilePath(tgtPath + "/doc02.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}

	@Test
	public void test_pptx_ppt01_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/ppt01.pptx");
		doc.setOutFilePath(tgtPath + "/ppt01.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}

	@Test
	public void test_ppt_pptsample_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/pptsample.ppt");
		doc.setOutFilePath(tgtPath + "/pptsample.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}

	@Test
	public void test_xlsx_excel01_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/excel01.xlsx");
		doc.setOutFilePath(tgtPath + "/excel01.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}
	
	@Test
	public void test_xls_excel02_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/excel02.xls");
		doc.setOutFilePath(tgtPath + "/excel02.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}

	@Test
	public void test_odt_UOMLSample_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/UOMLSample.odt");
		doc.setOutFilePath(tgtPath + "/UOMLSample.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}

//	@Test
	public void test_hwp_hwp01_to_pdf() {
		Document doc = new Document();
		doc.setInFilePath(resPath + "/hpw01.hpw");
		doc.setOutFilePath(tgtPath + "/hwp01.pdf");
		Converter converter = Parser.parse(doc);
		converter.convert();
	}	
}
