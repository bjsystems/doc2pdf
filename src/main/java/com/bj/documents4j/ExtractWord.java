package com.bj.documents4j;

import java.io.File;
import java.util.Properties;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.documents4j.throwables.ConverterException;

public class ExtractWord {
	public boolean extract() {
		boolean isSucc = true;
		Properties properties = new Properties();
//		properties.setProperty(FileRow.INPUT_NAME_PROPERTY_KEY, fileUpload.getClientFileName());
//		properties.setProperty(FileRow.SOURCE_FORMAT, sourceFormat.getModelObject().toString());
//		properties.setProperty(FileRow.TARGET_FORMAT, targetFormat.getModelObject().toString());
		
		long conversionDuration;
		try {
			conversionDuration = System.currentTimeMillis();
//			DemoApplication.get().getConverter()
//			.convert(newFile).as(sourceFormat.getModelObject())
//			.to(target, conductor).as(targetFormat.getModelObject())
//			.execute();
			File wordFile = new File("/demo.doc");
			File target = new File("/demo.pdf");
			File baseFolder = new File("/Users/baesunghan/Downloads");
			IConverter converter = LocalConverter.builder()
                    .baseFolder(baseFolder)
                    .workerPool(20, 25, 2, TimeUnit.SECONDS)
                    .processTimeout(5, TimeUnit.SECONDS)
                    .build();
			Future<Boolean> conversion = converter
			                                .convert(wordFile).as(DocumentType.MS_WORD)
			                                .to(target).as(DocumentType.PDF)
			                                .prioritizeWith(1000) // optional
			                                .schedule();
			
			conversionDuration = System.currentTimeMillis() - conversionDuration;
//			properties.setProperty(FileRow.CONVERSION_DURATION_PROPERTY_KEY, String.valueOf(conversionDuration));
//			writeProperties(properties, transactionFolder);
		} catch (ConverterException e) {
//			File[] transactionFiles = transactionFolder.listFiles();
//			if (transactionFiles != null) {
//				for (File file : transactionFiles) {
//					if (!file.delete()) {
//						LOGGER.warn("Could not delete transaction file {}", file);
//					}
//				}
//			}
//			if (!transactionFolder.delete()) {
//				LOGGER.warn("Could not delete transaction folder {}", transactionFolder);
//			}
//			/* other than this, this exception is already handled by the callback (FeedbackMessageConductor) */
		}
//		getFeedbackMessages().add(conductor.getFeedbackMessage());
		return isSucc;
	}

}
