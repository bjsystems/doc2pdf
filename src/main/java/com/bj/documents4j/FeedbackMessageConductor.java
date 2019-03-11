package com.bj.documents4j;

import java.io.File;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.documents4j.api.IFileConsumer;

public class FeedbackMessageConductor implements IFileConsumer {
	private static final Logger logger = LoggerFactory.getLogger(FeedbackMessageConductor.class);
    private final String inputName;

    private FeedbackMessageConductor(String inputName) {
        this.inputName = inputName;
    }

    @Override
    public void onComplete(File file) {
        String message = String.format("File '%s' was successfully converted.", inputName);
        logger.info(message);
    }

    @Override
    public void onCancel(File file) {
        String message = String.format("File conversion of '%s' was cancelled.", inputName);
        logger.error(message);
    }

    @Override
    public void onException(File file, Exception e) {
        String message = String.format("Could not convert file '%s'. Did you provide a valid MS Word file as input? [%s: %s]",
                inputName, e.getClass().getSimpleName(), e.getMessage());
        logger.error(message, e);
    }

}
