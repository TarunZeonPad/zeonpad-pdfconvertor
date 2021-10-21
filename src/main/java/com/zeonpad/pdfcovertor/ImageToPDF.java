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
import com.jacob.com.Variant;

public class ImageToPDF {

	private ActiveXComponent pdfApp;
	private boolean terminate;
	// Default Set to 5 minute.
	private int conversionTimeOut = 5;
	private String docFiledestination=null;

	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;

	/**
	 * @return the conversionTimeOut
	 */
	public int getConversionTimeOut() {
		return conversionTimeOut;
	}

	/**
	 * @param conversionTimeOut
	 *            the conversionTimeOut to set
	 */
	public void setConversionTimeOut(int conversionTimeOut) {
		this.conversionTimeOut = conversionTimeOut;
	}

	public void convert(String source, String destination) throws PDFException {
		
		PDFileValidation.validateFile(destination);
		ComThread.InitSTA();
		File fObj =null;
		try {

			pdfScheduler = Executors.newScheduledThreadPool(1);

			knightWatch = pdfScheduler.schedule(new Runnable() {

				public void run() {
					MSRunProcess.runPrepare(PDFCommandConst.KILL_WORD);
					pdfScheduler.shutdown();
					terminate = true;
				}
			}, conversionTimeOut, TimeUnit.MINUTES);

			UUID uuid = UUID.randomUUID();
			String docId = uuid.toString();
			docId = docId.replace("-", "") + "";

			docFiledestination = System.getProperty("java.io.tmpdir")
					+ docId + ".doc";
			fObj = new File(docFiledestination);

			if (!fObj.exists()) {
				fObj.createNewFile();
			}

			pdfApp = new ActiveXComponent(PDFCommandConst.MS_WORD);
			pdfApp.setProperty(PDFCommandConst.VISIBLE, false);
			ActiveXComponent oDocuments = pdfApp
					.getPropertyAsComponent(PDFCommandConst.DOCUMENTS);
			ActiveXComponent oDocument = oDocuments.invokeGetComponent("Open",
					new Variant(docFiledestination));

			ActiveXComponent oImage = oDocument
					.getPropertyAsComponent("InLineShapes");

			oImage.invoke("AddPicture", source);

			oDocument.invoke("SaveAs", docFiledestination);

			oDocument.invoke("SaveAs", destination, 17);

			//Dispatch.call(oDocument, PDFCommandConst.CLOSE, false);

		} catch (ComException pdfOxyExcp) {

			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			if(errorObj.getErrorCode() == 1012){
				errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MS_WORD);
			}
			throw new PDFException(errorObj.getErrorCode(),
					errorObj.getErrorDescription());
		} catch (Exception pdfOxyExcp) {
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			throw new PDFException(errorObj.getErrorCode(),
					errorObj.getErrorDescription());
		} finally {
			
			if(fObj != null && fObj.exists()){
				fObj.delete();
			}
			
			if (pdfApp != null && !terminate) {
				pdfApp.invoke(PDFCommandConst.QUIT,
						PDFNumConst.wordDoNotSaveChanges);
			}
			knightWatch.cancel(true);
			pdfScheduler.shutdown();

			ComThread.Release();
		}

	}
}
