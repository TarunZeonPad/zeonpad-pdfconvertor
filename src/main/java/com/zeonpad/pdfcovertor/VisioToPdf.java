/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   28-11-2017 
 */

package com.zeonpad.pdfcovertor;

import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class VisioToPdf {

	private ActiveXComponent pdfApp;
	private Dispatch pdfDispatch;
	
	//PDF fixed format
	private int visioToFixedFormatPDF=1;
	//Intended to be published online and printed
	private int visioDocExIntentPrint=1;
	//Default Prints all foreground pages.
	private int visioPrintOutRange;
	
	private int FromPage;
	
	private int ToPage;
	
	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;
	
	private boolean terminate;
	
	//True to render all colors as black to ensure that all shapes are visible in the exported drawing.  False to render colors normally. The default is False. 
	private boolean colorAsBlack;
	
	//Whether to include background pages in the exported file. The default is True.
	private boolean includeBackground=true;
	//Whether to include document properties in the exported file. The default is True .
	private boolean includeDocumentProperties=true;
	//Whether to include document structure tags to improve document accessibility. The default is True .
	private boolean includeStructureTags = true;
	//Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is False .
	private boolean nativeOfficePDF;
	// Default Set to 5 minute.
	private int conversionTimeOut = 5;
	
	
	/**
	 * @return the includeStructureTags
	 */
	public boolean isIncludeStructureTags() {
		return includeStructureTags;
	}

	/**
	 * @param includeStructureTags the includeStructureTags to set
	 */
	public void setIncludeStructureTags(boolean includeStructureTags) {
		this.includeStructureTags = includeStructureTags;
	}

	/**
	 * @return the includeDocumentProperties
	 */
	public boolean isIncludeDocumentProperties() {
		return includeDocumentProperties;
	}

	/**
	 * @param includeDocumentProperties the includeDocumentProperties to set
	 */
	public void setIncludeDocumentProperties(boolean includeDocumentProperties) {
		this.includeDocumentProperties = includeDocumentProperties;
	}

	/**
	 * @return the includeBackground
	 */
	public boolean isIncludeBackground() {
		return includeBackground;
	}

	/**
	 * @param includeBackground the includeBackground to set
	 */
	public void setIncludeBackground(boolean includeBackground) {
		this.includeBackground = includeBackground;
	}

	/**
	 * @return the colorAsBlack
	 */
	public boolean isColorAsBlack() {
		return colorAsBlack;
	}

	/**
	 * @param colorAsBlack the colorAsBlack to set
	 */
	public void setColorAsBlack(boolean colorAsBlack) {
		this.colorAsBlack = colorAsBlack;
	}

	/**
	 * @param fromPage the fromPage to set
	 */
	public void setFromPage(int fromPage) {
		FromPage = fromPage;
	}

	/**
	 * @param toPage the toPage to set
	 */
	public void setToPage(int toPage) {
		ToPage = toPage;
	}

	/**
	 * @return the visioPrintOutRange
	 */
	public int getVisioPrintOutRange() {
		return visioPrintOutRange;
	}

	/**
	 * @param visioPrintOutRange the visioPrintOutRange to set
	 */
	public void setVisioPrintOutRange(int visioPrintOutRange) {

		if (PDFVisioPrintOutRange.VISIO_PRINT_ALL == visioPrintOutRange
				|| PDFVisioPrintOutRange.VISIO_PRINT_FROM_TO == visioPrintOutRange
				|| PDFVisioPrintOutRange.VISIO_PRINT_CURRENT_PAGE == visioPrintOutRange
				|| PDFVisioPrintOutRange.VISIO_PRINT_SELECTION == visioPrintOutRange
				|| PDFVisioPrintOutRange.VISIO_PRINT_CURRENT_VIEW == visioPrintOutRange) {
			this.visioPrintOutRange = visioPrintOutRange;
		}

	}

	/**
	 * @return the nativeOfficePDF
	 */
	public boolean isNativeOfficePDF() {
		return nativeOfficePDF;
	}

	/**
	 * @param nativeOfficePDF the nativeOfficePDF to set
	 */
	public void setNativeOfficePDF(boolean nativeOfficePDF) {
		this.nativeOfficePDF = nativeOfficePDF;
	}

	/**
	 * @return the conversionTimeOut
	 */
	public int getConversionTimeOut() {
		return conversionTimeOut;
	}

	/**
	 * @param conversionTimeOut the conversionTimeOut to set
	 */
	public void setConversionTimeOut(int conversionTimeOut) {
		this.conversionTimeOut = conversionTimeOut;
	}

	public void convert(String source, String destination) throws PDFException{
	
		ComThread.InitSTA();
		try {
		
			pdfScheduler = Executors.newScheduledThreadPool(1);

			knightWatch = pdfScheduler.schedule(new Runnable() {

				public void run() {
					MSRunProcess.runPrepare(PDFCommandConst.KILL_VISIO);
					pdfScheduler.shutdown();
					terminate = true;
				}
			}, conversionTimeOut, TimeUnit.MINUTES);	
			
		pdfApp = new ActiveXComponent(PDFCommandConst.MS_VISIO);
		pdfApp.setProperty(PDFCommandConst.VISIBLE, new Variant(false));
		pdfDispatch = pdfApp.getProperty(PDFCommandConst.DOCUMENTS).toDispatch();
		
		Dispatch dispatchObj = Dispatch.call(pdfDispatch,
				PDFCommandConst.OPEN, source)
				.toDispatch();
		
			Object[] args = new Object[] { 
					new Variant(visioToFixedFormatPDF)//PDF fixed format
					,destination //File Path
					,new Variant(visioDocExIntentPrint)//Intended to be published online and printed
					,new Variant(visioPrintOutRange)//The range of document pages to be exported.  
					,new Variant(FromPage) 
					,new Variant(ToPage)
					,colorAsBlack
					,includeBackground
					,includeDocumentProperties
					,includeStructureTags
					,nativeOfficePDF
					};

		Dispatch.invoke(dispatchObj, PDFCommandConst.EXPORTASFIXEDFORMAT,
				Dispatch.Method, args, new int[1]);
		
		Dispatch.call(dispatchObj, "Close");

		} catch (ComException pdfOxyExcp) {
			destination=null;
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			if(errorObj.getErrorCode() == 1012){
				errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MS_VISIO);
			}
			throw new PDFException(errorObj.getErrorCode(),
					errorObj.getErrorDescription());
		} catch (Exception pdfOxyExcp) {
			destination=null;
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			throw new PDFException(errorObj.getErrorCode(),
					errorObj.getErrorDescription());
		} finally {
			if (pdfApp != null && !terminate) {
				pdfApp.invoke(PDFCommandConst.QUIT);
			}
			knightWatch.cancel(true);
			pdfScheduler.shutdown();
			ComThread.Release();
		}
		
		
	}
}
