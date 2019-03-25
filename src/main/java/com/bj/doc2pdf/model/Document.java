package com.bj.doc2pdf.model;

public class Document {
	private String Type;
	private String InFilePath;
	private String OutFilePath;
	private boolean showMessage = false;
	
	public String getInFilePath() {
		return InFilePath;
	}
	public void setInFilePath(String inFilePath) {
		InFilePath = inFilePath;
	}
	public String getOutFilePath() {
		return OutFilePath;
	}
	public void setOutFilePath(String outFilePath) {
		OutFilePath = outFilePath;
	}
	public String getType() {
		return Type;
	}
	public void setType(String type) {
		Type = type;
	}
	public boolean isShowMessage() {
		return showMessage;
	}
	public void setShowMessage(boolean showMessage) {
		this.showMessage = showMessage;
	}

}
