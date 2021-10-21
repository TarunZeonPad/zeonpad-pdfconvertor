/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;



public class PDFErrorGenerator {

	public static PDFErrorPojo generateErrorMsg(String errDescription,boolean termiate){
		
		PDFErrorPojo errorObj = new PDFErrorPojo();
		
		if(termiate){
			errorObj.setErrorCode(PDFErrorMessages.PDF_CV_PROCESS_TERMINATED_CODE);
			errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_PROCESS_TERMINATED);
			return errorObj;
		}
		
		String errorDescp = "";
		if(errDescription != null){
			errorDescp = errDescription.toLowerCase();
		}
		
		if(errorDescp.contains("password")){
			errorObj.setErrorCode(PDFErrorMessages.PDF_CV_PASSWORD_PROTECTED_CODE);
			errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_PASSWORD_PROTECTED);
			return errorObj;
			
		}
		
		if(errorDescp.contains("open")){
			errorObj.setErrorCode(PDFErrorMessages.PDF_CV_UNABLE_OPEN_CODE);
			errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_UNABLE_OPEN);
			return errorObj;
			
		}
		
		if(errorDescp.contains("exception")){
			errorObj.setErrorCode(PDFErrorMessages.PDF_CV_UNABLE_OPEN_CODE);
			errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_UNABLE_OPEN);
			return errorObj;
			
		}
		
		if(errorDescp.contains("read-only")){
			errorObj.setErrorCode(PDFErrorMessages.PDF_CV_PERMISSION_ERR_CODE);
			errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_PERMISSION_ERR);
			return errorObj;
			
		}
		
		if(errorDescp.contains("can't get object clsid from progid")){
			errorObj.setErrorCode(PDFErrorMessages.PDF_CV_NO_MSOFFICE_ERR_CODE);
			errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MSOFFICE_ERR_DESC);
			return errorObj;
		}
			
		errorObj.setErrorCode(PDFErrorMessages.PDF_CV_CONVERSION_ERR_CODE);
		errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_CONVERSION_ERR);
		
		return errorObj;
	}
	
}
