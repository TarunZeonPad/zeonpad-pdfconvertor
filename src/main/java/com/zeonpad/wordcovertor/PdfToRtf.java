package com.zeonpad.wordcovertor;

import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.zeonpad.pdfcovertor.MSRunProcess;
import com.zeonpad.pdfcovertor.PDFCommandConst;
import com.zeonpad.pdfcovertor.PDFErrorGenerator;
import com.zeonpad.pdfcovertor.PDFErrorMessages;
import com.zeonpad.pdfcovertor.PDFErrorPojo;
import com.zeonpad.pdfcovertor.PDFException;
import com.zeonpad.pdfcovertor.PDFNumConst;
import com.zeonpad.pdfcovertor.PDFileValidation;

public class PdfToRtf {

	private ActiveXComponent pdfApp;
	private Dispatch pdfDispatch;

	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;
	private boolean terminate;
	private String documentPassword = "ZeonPassword";

	// Default Set to 7 minute.
	private int conversionTimeOut = 7;
	
	/**
	 * @param documentPassword
	 *            the documentPassword to set
	 */
	public void setDocumentPassword(String documentPassword) {
		this.documentPassword = documentPassword;
	}
	
	/**
	 * @return the setConversionTimeOut
	 */
	public int getConversionTimeOut() {
		return conversionTimeOut;
	}

	/**
	 * @param setConversionTimeOut
	 *            the setConversionTimeOut to set
	 */
	public void setConversionTimeOut(int conversionTimeOut) {
		this.conversionTimeOut = conversionTimeOut;
	}


	public void convert(String source, String destination) throws PDFException {

		PDFileValidation.validateOutputRtfFile(destination);
		ComThread.InitSTA();

		try {

			pdfScheduler = Executors.newScheduledThreadPool(1);

			knightWatch = pdfScheduler.schedule(new Runnable() {

				public void run() {
					MSRunProcess.runPrepare(PDFCommandConst.KILL_WORD);
					pdfScheduler.shutdown();
					terminate = true;
				}
			}, conversionTimeOut, TimeUnit.MINUTES);

			pdfApp = new ActiveXComponent(PDFCommandConst.MS_WORD);
			pdfApp.setProperty(PDFCommandConst.VISIBLE, false);// For performing
																// all action on
																// backend.
			pdfApp.setProperty(PDFCommandConst.DISPLAYALERTS, PDFCommandConst.DISPLAY_ALERTS);
			pdfDispatch = pdfApp.getProperty(PDFCommandConst.DOCUMENTS).toDispatch();

			Dispatch doc = Dispatch
					.invoke(pdfDispatch, PDFCommandConst.OPEN, Dispatch.Method,
							new Object[] { source, new Variant(false), new Variant(false), new Variant(false), documentPassword, documentPassword,
									new Variant(false), "", "", PDFNumConst.wdOpenFormatAuto },
							new int[1])
					.toDispatch();

			Dispatch.invoke(doc, PDFCommandConst.SAVEAS, Dispatch.Method,
					new Object[] { destination, new Variant(PDFNumConst.wdFormatXMLDocument), new Variant(false), "",
							new Variant(true), "", new Variant(false), new Variant(false), new Variant(false),
							new Variant(false), new Variant(false) },
					new int[1]);
			
			Variant f = new Variant(false);
			Dispatch.call(doc, PDFCommandConst.CLOSE, f);

		} catch (ComException pdfOxyExcp) {

			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(pdfOxyExcp.getMessage(), terminate);
			if (errorObj.getErrorCode() == 1012) {
				errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MS_WORD);
			}
			throw new PDFException(errorObj.getErrorCode(), errorObj.getErrorDescription());
		} catch (Exception pdfOxyExcp) {
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(pdfOxyExcp.getMessage(), terminate);
			throw new PDFException(errorObj.getErrorCode(), errorObj.getErrorDescription());
		} finally {
			if (pdfApp != null && !terminate) {
				pdfApp.invoke(PDFCommandConst.QUIT, PDFNumConst.wordDoNotSaveChanges);
			}
			knightWatch.cancel(true);
			pdfScheduler.shutdown();

			ComThread.Release();
		}

	}
}
