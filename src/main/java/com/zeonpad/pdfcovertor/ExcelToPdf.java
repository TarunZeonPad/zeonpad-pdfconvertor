/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   20-11-2017 
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

public class ExcelToPdf {
	
	
	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;

	ActiveXComponent pdfApp;
	Dispatch pdfDispatch;
	int pdfExportQuality;
	boolean includeDocProps;
	boolean ignorePrintAreas;
	boolean printAllSheets=true;
	int from = 1;
	int to = 1;
	private boolean terminate;
	// Default Password set to ZeonpadTesting
	private String documentPassword="ZeonpadTesting";
	//Default Set to 5 minute.
	private int conversionTimeOut=5;
	
	/**
	 * @return the pdfExportQuality
	 */
	public int getPdfExportQuality() {
		return pdfExportQuality;
	}

	/**
	 * @param pdfExportQuality
	 *            the pdfExportQuality to set
	 */
	public void setPdfExportQuality(int pdfExportQuality) {
		if (pdfExportQuality == PDFXlFixedFormatQuality.PDF_EXPORT_QUALITY_STANDARD
				|| pdfExportQuality == PDFXlFixedFormatQuality.PDF_EXPORT_QUALITY_MINIMUM) {
			this.pdfExportQuality = pdfExportQuality;
		}

	}

	/**
	 * @return the includeDocProps
	 */
	public boolean isIncludeDocProps() {
		return includeDocProps;
	}

	/**
	 * @param includeDocProps
	 *            the includeDocProps to set
	 */
	public void setIncludeDocProps(boolean includeDocProps) {
		this.includeDocProps = includeDocProps;
	}

	/**
	 * @return the ignorePrintAreas
	 */
	public boolean isIgnorePrintAreas() {
		return ignorePrintAreas;
	}

	/**
	 * @param ignorePrintAreas
	 *            the ignorePrintAreas to set
	 */
	public void setIgnorePrintAreas(boolean ignorePrintAreas) {
		this.ignorePrintAreas = ignorePrintAreas;
	}

	/**
	 * @return the printAllSheets
	 */
	public boolean isPrintAllSheets() {
		return printAllSheets;
	}

	/**
	 * @param printAllSheets
	 *            the printAllSheets to set
	 */
	public void setPrintAllSheets(boolean printAllSheets) {
		this.printAllSheets = printAllSheets;
	}

	/**
	 * @return the from
	 */
	public int getFrom() {
		return from;
	}

	/**
	 * @param from
	 *            the from to set
	 */
	public void setFrom(int from) {
		this.from = from;
	}

	/**
	 * @return the to
	 */
	public int getTo() {
		return to;
	}

	/**
	 * @param to
	 *            the to to set
	 */
	public void setTo(int to) {
		this.to = to;
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

	/**
	 * @param documentPassword the documentPassword to set
	 */
	public void setDocumentPassword(String documentPassword) {
		this.documentPassword = documentPassword;
	}

	public void convert(String source, String destination) throws PDFException {

		PDFileValidation.validateFile(destination);
		ComThread.InitSTA();
		try {

			

			if(!printAllSheets){
				if(from <=0){
					throw new PDFException(PDFErrorMessages.PDF_CV_FROMSHEET_ERR_CODE,
							PDFErrorMessages.PDF_CV_FROMSHEET_ERR);
				}
				if(to <=0){
					throw new PDFException(PDFErrorMessages.PDF_CV_TOSHEET_ERR_CODE,
							PDFErrorMessages.PDF_CV_TOSHEET_ERR);
				}		
			}
			
			pdfScheduler = Executors.newScheduledThreadPool(1);

			knightWatch = pdfScheduler.schedule(new Runnable() {

				public void run() {
					MSRunProcess.runPrepare(PDFCommandConst.KILL_EXCEL);
					pdfScheduler.shutdown();
					terminate = true;
				}
			}, conversionTimeOut, TimeUnit.MINUTES);

			
			
			pdfApp = new ActiveXComponent(PDFCommandConst.MS_EXCEL);
			pdfApp.setProperty(PDFCommandConst.VISIBLE, false);// For performing
																// all action on
																// backend.
			pdfApp.setProperty(PDFCommandConst.DISPLAYALERTS, PDFCommandConst.DISPLAY_ALERTS);
			pdfDispatch = pdfApp.getProperty(PDFCommandConst.WORKBOOKS)
					.toDispatch();
			
			Dispatch dispatchObj = Dispatch.call(pdfDispatch,
					PDFCommandConst.OPEN, source, false, true,true,documentPassword).toDispatch();
			
			Object[] args =null;
			
			if(printAllSheets){
				args = new Object[] {
						new Integer(PDFNumConst.excelToFormatPDF), 
						destination,
						new Integer(pdfExportQuality),
						new Boolean(includeDocProps),
						new Boolean(ignorePrintAreas) 
						};

			}else{
				args = new Object[] {
						new Integer(PDFNumConst.excelToFormatPDF), 
						destination,
						new Integer(pdfExportQuality),
						new Boolean(includeDocProps),
						new Boolean(ignorePrintAreas), 
						new Integer(from),
						new Integer(to) };

			}
			
			
			Dispatch.invoke(dispatchObj, PDFCommandConst.EXPORTASFIXEDFORMAT,
					Dispatch.Method, args, new int[1]);

			Dispatch.call(dispatchObj, "Close", false);

		} catch (ComException pdfOxyExcp) {
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			if(errorObj.getErrorCode() == 1012){
				errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MS_EXCEL);
			}
			throw new PDFException(errorObj.getErrorCode(),
					errorObj.getErrorDescription());
		} catch (Exception pdfOxyExcp) {
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			throw new PDFException(errorObj.getErrorCode(),
					errorObj.getErrorDescription());
		}  finally {
			if (pdfApp != null && !terminate) {
				pdfApp.invoke(PDFCommandConst.QUIT, new Variant[] {});
			}
			
			knightWatch.cancel(true);
			pdfScheduler.shutdown();

		}
		ComThread.Release();
	}

}
