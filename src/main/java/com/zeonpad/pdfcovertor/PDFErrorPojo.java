/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;

public class PDFErrorPojo {

	private int errorCode;
	private String errorDescription;
	/**
	 * @return the errorCode
	 */
	public int getErrorCode() {
		return errorCode;
	}
	/**
	 * @param errorCode the errorCode to set
	 */
	public void setErrorCode(int errorCode) {
		this.errorCode = errorCode;
	}
	/**
	 * @return the errorDescription
	 */
	public String getErrorDescription() {
		return errorDescription;
	}
	/**
	 * @param errorDescription the errorDescription to set
	 */
	public void setErrorDescription(String errorDescription) {
		this.errorDescription = errorDescription;
	}
}
