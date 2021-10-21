/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   22-11-2017 
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

public class PptToPdf {

	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;
	private ActiveXComponent pdfApp;
	private Dispatch pdfDispatch;
	private Dispatch dispatchObj;
	private boolean terminate;
	
	//Default Set to 5 minute.
	private int conversionTimeOut=5;

	/**
	 * @return the setConversionTimeOut
	 */
	public int getConversionTimeOut() {
		return conversionTimeOut;
	}

	/**
	 * @param setConversionTimeOut the setConversionTimeOut to set
	 */
	public void setConversionTimeOut(int conversionTimeOut) {
		this.conversionTimeOut = conversionTimeOut;
	}

	public void convert(String source, String destination) throws PDFException {
		PDFileValidation.validateFile(destination);
		ComThread.InitSTA();
		try {

			pdfScheduler = Executors.newScheduledThreadPool(1);

			knightWatch = pdfScheduler.schedule(new Runnable() {

				public void run() {
					MSRunProcess.runPrepare(PDFCommandConst.KILL_PPT);
					pdfScheduler.shutdown();
					terminate = true;
				}
			}, conversionTimeOut, TimeUnit.MINUTES);

			pdfApp = new ActiveXComponent(PDFCommandConst.MS_PPT);
			
			pdfApp.setProperty(PDFCommandConst.DISPLAYALERTS, PDFCommandConst.DISPLAY_ALERTS);
			pdfDispatch = pdfApp.getProperty(PDFCommandConst.PRESENTATIONS)
					.toDispatch();
			dispatchObj = Dispatch.call(pdfDispatch, PDFCommandConst.OPEN,
					source, false, false, false).toDispatch();
			
			Object[] args = new Object[] { destination,
					new Variant(PDFNumConst.pptToFormatPDF) };

			Dispatch.invoke(dispatchObj, PDFCommandConst.SAVEAS,
					Dispatch.Method, args, new int[1]);

			Dispatch.call(dispatchObj, PDFCommandConst.CLOSE);

		} catch (ComException pdfOxyExcp) {
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			if(errorObj.getErrorCode() == 1012){
				errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MS_PPT);
			}
			throw new PDFException(errorObj.getErrorCode(),
					errorObj.getErrorDescription());
		} catch (Exception pdfOxyExcp) {
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
