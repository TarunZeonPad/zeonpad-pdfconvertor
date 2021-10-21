/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;

public abstract interface PDFErrorMessages {

	public static final int PDF_CV_PROCESS_TERMINATED_CODE = 1000;

	public static final String PDF_CV_PROCESS_TERMINATED = "Host application terminated during conversion because unable to open input document either document is password protected or it has some popup.";

	public static final int PDF_CV_UNABLE_OPEN_CODE = 1001;

	public static final String PDF_CV_UNABLE_OPEN = "Unable to open input document.";

	public static final int PDF_CV_END_VAL_NDF_CODE = 1002;

	public static final String PDF_CV_END_VAL_NDF = "End Page value is not defined.";
	
	public static final int PDF_CV_PASSWORD_PROTECTED_CODE  = 1003;
	
	public static final String PDF_CV_PASSWORD_PROTECTED  = "The document is password protected.";
	
	public static final int PDF_CV_UNKNOWN_ERR_CODE  = 1004;
	
	public static final String PDF_CV_UNKNOWN_ERR  = "Unknown error while coverting document to pdf";
	
	public static final int PDF_CV_FROMSHEET_ERR_CODE  = 1005;
	
	public static final String PDF_CV_FROMSHEET_ERR = "fromSheet value is not within the Page Range,Enter a Number between 1 and 32766.";

	public static final int PDF_CV_TOSHEET_ERR_CODE  = 1006;
	
	public static final String PDF_CV_TOSHEET_ERR = "toSheet value is not within the Page Range,Enter a Number between 1 and 32766.";

	public static final int PDF_CV_CONVERSION_ERR_CODE  = 1007;
	
	public static final String PDF_CV_CONVERSION_ERR = "Conversion error.";
	
	public static final int PDF_CV_PERMISSION_ERR_CODE  = 1008;
	
	public static final String PDF_CV_PERMISSION_ERR = "Unable to create output file because of permission issue.";
	
	public static final int PDF_CV_NO_EXTENSION_ERR_CODE  = 1009;
	
	public static final String PDF_CV_NO_EXTENSION_ERR_DESC  = "Output file name does not have pdf extension.";
	
	public static final int PDF_CV_CREATE_TEMPFILE_ERR_CODE  = 1010;
	
	public static final String PDF_CV_CREATE_TEMPFILE_ERR_DESC  = "Unable to create temporary file because of permission issue.";
	
	public static final int PDF_CV_NO_EXTENSION_INPUT_ERR_CODE  = 1011;
	
	public static final String PDF_CV_NO_EXTENSION_INPUT_ERR_DESC  = "Input file name does not have any extension.";
	
	public static final int PDF_CV_NO_MSOFFICE_ERR_CODE  = 1012;
	
	public static final String PDF_CV_NO_MSOFFICE_ERR_DESC  = "MS OFFICE is not installed";
	
	public static final int PDF_DOC_NO_EXTENSION_ERR_CODE  = 1013;
	
	public static final String PDF_DOC_NO_EXTENSION_ERR_DESC  = "Output file name does not have doc/docx extension.";
	
	public static final int PDF_RTF_NO_EXTENSION_ERR_CODE  = 1014;
	
	public static final String PDF_RTF_NO_EXTENSION_ERR_DESC  = "Output file name does not have rtf extension.";
	
	public static final int PDF_HTML_NO_EXTENSION_ERR_CODE  = 1014;
	
	public static final String PDF_HTML_NO_EXTENSION_ERR_DESC  = "Output file name does not have html extension.";
	
	public static final int PDF_TXT_NO_EXTENSION_ERR_CODE  = 1015;
	
	public static final String PDF_TXT_NO_EXTENSION_ERR_DESC  = "Output file name does not have txt extension.";
	
	public static final String PDF_CV_NO_MS_WORD ="Microsoft Word is not installed, Please install and try again.";
	
	public static final String PDF_CV_NO_MS_EXCEL ="Microsoft Excel is not installed, Please install and try again.";
	
	public static final String PDF_CV_NO_MS_PPT ="Microsoft Powerpoint is not installed, Please install and try again.";
	
	public static final String PDF_CV_NO_MS_PUB ="Microsoft Publisher is not installed, Please install and try again.";
	
	public static final String PDF_CV_NO_MS_MSG ="Microsoft Outlook is not installed, Please install and try again.";
	
	public static final String PDF_CV_NO_MS_IMG ="Microsoft Word is not installed, Please install and try again.";
	
	public static final String PDF_CV_NO_MS_VISIO ="Microsoft Visio is not installed, Please install and try again.";
	
	

}
