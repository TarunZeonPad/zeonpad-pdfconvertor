/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;

import java.io.File;
import java.util.UUID;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class OutlookToPdf {

	ActiveXComponent pdfApp;
	Dispatch pdfDispatch;
	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;
	private boolean terminate;

	// Default Set to 5 minute.
	private int conversionTimeOut = 5;

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

		String inputFile = convertToDoc(source);
		if (inputFile != null) {
			WordToPdf wordToPdf = new WordToPdf();
			wordToPdf.convert(inputFile, destination);
			try{
				if(inputFile != null && !"".equals(inputFile)){
					File file = new File(inputFile);
					System.out.println(file);
					boolean delete  = file.delete();
					System.out.println(delete);
				}
			}catch(Exception e){
				
			}
			
			
		} else {
			throw new PDFException(PDFErrorMessages.PDF_CV_UNKNOWN_ERR_CODE,
					PDFErrorMessages.PDF_CV_UNKNOWN_ERR);
		}

	}

	private String convertToDoc(String source) throws PDFException {

		ComThread.InitSTA();
		String destination = null;
		try {

			pdfScheduler = Executors.newScheduledThreadPool(1);

			knightWatch = pdfScheduler.schedule(new Runnable() {

				public void run() {
					MSRunProcess.runPrepare(PDFCommandConst.KILL_OUTLOOK);
					pdfScheduler.shutdown();
					terminate = true;
				}
			}, conversionTimeOut, TimeUnit.MINUTES);

			UUID uuid = UUID.randomUUID();
			String docId = uuid.toString();
			docId = docId.replace("-", "") + "";

			destination = System.getProperty("java.io.tmpdir") + docId+".doc";

			//destination = "E://TempFolder//" + docId+".doc";
				
			pdfApp = new ActiveXComponent(PDFCommandConst.MS_OUTLOOK);

			pdfDispatch = pdfApp.getProperty(PDFCommandConst.SESSION).toDispatch();

			Dispatch dispatchObj = Dispatch.call(pdfDispatch, PDFCommandConst.OPENSHAREDITEM,
					source).toDispatch();

			Object[] args = new Object[] { destination, // Output file name
					new Variant(4) // PbFixedFormatType
			};

			 Dispatch.invoke(dispatchObj, PDFCommandConst.SAVEAS,
					Dispatch.Method, args, new int[1]);
			
			Dispatch.call(dispatchObj, PDFCommandConst.CLOSE, false);

		} catch (ComException pdfOxyExcp) {
			destination=null;
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			if(errorObj.getErrorCode() == 1012){
				errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MS_MSG);
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
		return destination;
	}
}
