/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;

import java.io.File;

public class PDFileValidation {
	

	public static void validateFile(String destination) throws PDFException
	{
	
		File f = new File(destination);		

		String fileName = f.getName();
		
		try{
			String fileExtension = fileName.substring(fileName.lastIndexOf("."));
			if(!fileExtension.contains(".pdf")){
				throw new PDFException(PDFErrorMessages.PDF_CV_NO_EXTENSION_ERR_CODE,
						PDFErrorMessages.PDF_CV_NO_EXTENSION_ERR_DESC);
			}
			
		}catch(Exception e){
			throw new PDFException(PDFErrorMessages.PDF_CV_NO_EXTENSION_ERR_CODE,
					PDFErrorMessages.PDF_CV_NO_EXTENSION_ERR_DESC);
		}
		
		
		/*if(!f.canWrite()){
			throw new PDFException(PDFErrorMessages.PDF_CV_PERMISSION_ERR_CODE,
					PDFErrorMessages.PDF_CV_PERMISSION_ERR);
		}*/
		
	}
	
	public static void validateInputFile(String source) throws PDFException
	{
	
		File f = new File(source);		

		String fileName = f.getName();
		
		try{
			String fileExtension = fileName.substring(fileName.lastIndexOf("."));
		}catch(Exception e){
			throw new PDFException(PDFErrorMessages.PDF_CV_NO_EXTENSION_INPUT_ERR_CODE,
					PDFErrorMessages.PDF_CV_NO_EXTENSION_INPUT_ERR_DESC);
		}
		
	}
	
	
	public static void validateOutputDocFile(String destination) throws PDFException
	{
	
		File f = new File(destination);		

		String fileName = f.getName();
		
		try{
			String fileExtension = fileName.substring(fileName.lastIndexOf("."));
			//System.out.println(fileExtension);
			if(fileExtension != null && !fileExtension.startsWith(".do")) {
				throw new PDFException(PDFErrorMessages.PDF_DOC_NO_EXTENSION_ERR_CODE,
						PDFErrorMessages.PDF_DOC_NO_EXTENSION_ERR_DESC);
			}
		}catch(Exception e){
			throw new PDFException(PDFErrorMessages.PDF_DOC_NO_EXTENSION_ERR_CODE,
					PDFErrorMessages.PDF_DOC_NO_EXTENSION_ERR_DESC);
		}
		
	}
	
	public static void validateOutputRtfFile(String destination) throws PDFException
	{
	
		File f = new File(destination);		

		String fileName = f.getName();
		
		try{
			String fileExtension = fileName.substring(fileName.lastIndexOf("."));
			//System.out.println(fileExtension);
			if(fileExtension != null && !fileExtension.startsWith(".rtf")) {
				throw new PDFException(PDFErrorMessages.PDF_RTF_NO_EXTENSION_ERR_CODE,
						PDFErrorMessages.PDF_RTF_NO_EXTENSION_ERR_DESC);
			}
		}catch(Exception e){
			throw new PDFException(PDFErrorMessages.PDF_RTF_NO_EXTENSION_ERR_CODE,
					PDFErrorMessages.PDF_RTF_NO_EXTENSION_ERR_DESC);
		}
		
	}
	
	public static void validateOutputHtmlFile(String destination) throws PDFException
	{
	
		File f = new File(destination);		

		String fileName = f.getName();
		
		try{
			String fileExtension = fileName.substring(fileName.lastIndexOf("."));
			//System.out.println(fileExtension);
			if(fileExtension != null && !fileExtension.startsWith(".htm")) {
				throw new PDFException(PDFErrorMessages.PDF_HTML_NO_EXTENSION_ERR_CODE,
						PDFErrorMessages.PDF_HTML_NO_EXTENSION_ERR_DESC);
			}
		}catch(Exception e){
			throw new PDFException(PDFErrorMessages.PDF_HTML_NO_EXTENSION_ERR_CODE,
					PDFErrorMessages.PDF_HTML_NO_EXTENSION_ERR_DESC);
		}
		
	}
	
	public static void validateOutputTxtFile(String destination) throws PDFException
	{
	
		File f = new File(destination);		

		String fileName = f.getName();
		
		try{
			String fileExtension = fileName.substring(fileName.lastIndexOf("."));
			//System.out.println(fileExtension);
			if(fileExtension != null && !fileExtension.startsWith(".txt")) {
				throw new PDFException(PDFErrorMessages.PDF_TXT_NO_EXTENSION_ERR_CODE,
						PDFErrorMessages.PDF_TXT_NO_EXTENSION_ERR_DESC);
			}
		}catch(Exception e){
			throw new PDFException(PDFErrorMessages.PDF_TXT_NO_EXTENSION_ERR_CODE,
					PDFErrorMessages.PDF_TXT_NO_EXTENSION_ERR_DESC);
		}
		
	}

	
	
	public static void validateTempFile(String destination) throws PDFException
	{
	
		/*System.out.println(destination);
		File f = new File("E://TempFolder");		
		System.out.println(f.canWrite());
		if(!f.canWrite()){
			throw new PDFException(PDFErrorMessages.PDF_CV_CREATE_TEMPFILE_ERR_CODE,
					PDFErrorMessages.PDF_CV_CREATE_TEMPFILE_ERR_DESC);
		}*/
		
	}

}
