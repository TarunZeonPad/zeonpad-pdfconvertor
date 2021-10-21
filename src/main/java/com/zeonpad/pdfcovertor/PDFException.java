/**
* @author  Tarun Kumar 
* @version 1.0
* @since   16-11-2017 
*/

package com.zeonpad.pdfcovertor;

@SuppressWarnings("serial")
public class PDFException extends Exception {
	
	protected int errorCode;

	public PDFException() {
	}

	public PDFException(int errorCode, String errorDescription) {
		super(errorDescription);
		this.errorCode = errorCode;
	}

	public String getMessage() {

		return super.getMessage();
	}

	public int getErrorCode() {
		return this.errorCode;
	}
}
