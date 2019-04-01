package com.bj.doc2pdf.convert;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.bj.doc2pdf.model.Document;

public class Parser {
	private static final Logger logger = LoggerFactory.getLogger(Parser.class);
	public enum DOC_TYPE {
		DOC,
		DOCX,
		PPT,
		PPTX,
		ODT
	}
	
	public static Converter parse(Document doc) {
		Converter converter = null;
		if (StringUtils.isEmpty(doc.getInFilePath())) {
			logger.error("Not found input file!!!");
			throw new IllegalArgumentException();
		}

		if (StringUtils.isEmpty(doc.getOutFilePath())) {
			doc.setOutFilePath(changeExtensionToPDF(doc.getInFilePath()));
		}

		String lowerCaseInPath = doc.getInFilePath().toLowerCase();

		try {
			InputStream inStream = getInFileStream(doc.getInFilePath());
			OutputStream outStream = getOutFileStream(doc.getOutFilePath());

			if (StringUtils.isEmpty(doc.getType())) {
				if (lowerCaseInPath.endsWith("doc")) {
					converter = new DocToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
				} else if (lowerCaseInPath.endsWith("docx")) {
					converter = new DocxToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
				} else if (lowerCaseInPath.endsWith("ppt")) {
					converter = new PptToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
				} else if (lowerCaseInPath.endsWith("pptx")) {
					converter = new PptxToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
				} else if (lowerCaseInPath.endsWith("xls")) {
					converter = new XlsToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
				} else if (lowerCaseInPath.endsWith("xlsx")) {
					converter = new XlsxToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
				} else if (lowerCaseInPath.endsWith("odt")) {
					converter = new OdtToPDF(inStream, outStream, doc.isShowMessage(), true);
				} else {
					converter = null;
				}
//			} else {
//				switch (doc.getType()) {
//				case DOC:
//					converter = new DocToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
//					break;
//				case DOCX:
//					converter = new DocxToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
//					break;
//				case PPT:
//					converter = new PptToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
//					break;
//				case PPTX:
//					converter = new PptxToPDFConverter(inStream, outStream, doc.isShowMessage(), true);
//					break;
//				case ODT:
//					converter = new OdtToPDF(inStream, outStream, doc.isShowMessage(), true);
//					break;
//				default:
//					converter = null;
//					break;
//				}
			}
		} catch (IOException e) {
			logger.error("File not found!!!");
			throw new RuntimeException("Exception: ", e);
		}

		return converter;
	}

	//From http://stackoverflow.com/questions/941272/how-do-i-trim-a-file-extension-from-a-string-in-java
	private static String changeExtensionToPDF(String originalPath) {

//		String separator = System.getProperty("file.separator");
		String filename = originalPath;

//		// Remove the path upto the filename.
//		int lastSeparatorIndex = originalPath.lastIndexOf(separator);
//		if (lastSeparatorIndex == -1) {
//			filename = originalPath;
//		} else {
//			filename = originalPath.substring(lastSeparatorIndex + 1);
//		}

		// Remove the extension.
		int extensionIndex = filename.lastIndexOf(".");

		String removedExtension;
		if (extensionIndex == -1){
			removedExtension =  filename;
		} else {
			removedExtension =  filename.substring(0, extensionIndex);
		}
		String addPDFExtension = removedExtension + ".pdf";

		return addPDFExtension;
	}
	
	protected static InputStream getInFileStream(String inputFilePath) throws FileNotFoundException{
		File inFile = new File(inputFilePath);
		FileInputStream iStream = new FileInputStream(inFile);
		return iStream;
	}
	
	protected static OutputStream getOutFileStream(String outputFilePath) throws IOException{
		File outFile = new File(outputFilePath);
		
		try{
			//Make all directories up to specified
			outFile.getParentFile().mkdirs();
		} catch (NullPointerException e){
			//Ignore error since it means not parent directories
		}
		
		outFile.createNewFile();
		FileOutputStream oStream = new FileOutputStream(outFile);
		return oStream;
	}


}
