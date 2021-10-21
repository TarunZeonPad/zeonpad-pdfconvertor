/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   30-11-2017 
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

public class PublisherToPdf {

	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;
	private ActiveXComponent pdfApp;
	private Dispatch pdfDispatch;
	private boolean includeDocProps;
	private int from=1;
	private int to=-1;
	private int copies=2;
	private boolean collate=false;
	private boolean docStructureTags = true;
	private boolean bitmapMissingFonts = true;
	private boolean useISO19005_1 = false;
	private boolean terminate;
	
	//Default Set to 5 minute.
	private int conversionTimeOut=5;
	
	/**
	 * @return the from
	 */
	public int getFrom() {
		return from;
	}

	/**
	 * @param from the from to set
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
	 * @param to the to to set
	 */
	public void setTo(int to) {
		this.to = to;
	}

	/**
	 * @return the includeDocProps
	 */
	public boolean isIncludeDocProps() {
		return includeDocProps;
	}

	/**
	 * @param includeDocProps the includeDocProps to set
	 */
	public void setIncludeDocProps(boolean includeDocProps) {
		this.includeDocProps = includeDocProps;
	}

	/**
	 * @return the copies
	 */
	public int getCopies() {
		return copies;
	}

	/**
	 * @param copies the copies to set
	 */
	public void setCopies(int copies) {
		this.copies = copies;
	}

	/**
	 * @return the collate
	 */
	public boolean isCollate() {
		return collate;
	}

	/**
	 * @param collate the collate to set
	 */
	public void setCollate(boolean collate) {
		this.collate = collate;
	}

	/**
	 * @return the docStructureTags
	 */
	public boolean isDocStructureTags() {
		return docStructureTags;
	}

	/**
	 * @param docStructureTags the docStructureTags to set
	 */
	public void setDocStructureTags(boolean docStructureTags) {
		this.docStructureTags = docStructureTags;
	}

	/**
	 * @return the bitmapMissingFonts
	 */
	public boolean isBitmapMissingFonts() {
		return bitmapMissingFonts;
	}

	/**
	 * @param bitmapMissingFonts the bitmapMissingFonts to set
	 */
	public void setBitmapMissingFonts(boolean bitmapMissingFonts) {
		this.bitmapMissingFonts = bitmapMissingFonts;
	}

	/**
	 * @return the useISO19005_1
	 */
	public boolean isUseISO19005_1() {
		return useISO19005_1;
	}

	/**
	 * @param useISO19005_1 the useISO19005_1 to set
	 */
	public void setUseISO19005_1(boolean useISO19005_1) {
		this.useISO19005_1 = useISO19005_1;
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

	public void convert(String source, String destination) throws PDFException {
		
		PDFileValidation.validateFile(destination);
		ComThread.InitSTA();
		try {
			
			pdfScheduler = Executors.newScheduledThreadPool(1);

			knightWatch = pdfScheduler.schedule(new Runnable() {

				public void run() {
					MSRunProcess.runPrepare(PDFCommandConst.KILL_PUB);
					pdfScheduler.shutdown();
					terminate = true;
				}
			}, conversionTimeOut, TimeUnit.MINUTES);

			pdfApp = new ActiveXComponent(PDFCommandConst.MS_PUBLISHER);
			pdfDispatch = pdfApp.getObject();

			Dispatch dispatchObj = Dispatch.call(pdfDispatch,
					PDFCommandConst.OPEN, source, true, true, PDFNumConst.publisherDoNotSaveChanges)
					.toDispatch();

			Object[] args = new Object[] {
					new Variant(PDFNumConst.publisherToFormatPDF) //PbFixedFormatType
					,destination  //Output file name
					,new Variant(3)  //PbFixedFormatIntent
					,includeDocProps //save the document properties with the PDF file
					,PDFNumConst.colorDownsampleTarget //Measured in dots per inch.
					,PDFNumConst.colorDownsampleThreshold //The threshold at or above which an image is down-sampled to the ColorDownsample target level.
					,PDFNumConst.oneBitDownsampleTarget //The target for down-sampling of one-bit images.
					,PDFNumConst.OneBitDownsampleThreshold //The threshold at or above which an image is down-sampled to the OneBitDownsample target level.
					,from
					,to
					,copies //The number of copies.
					,collate
					,0//PbPrintStyle 
					,docStructureTags
					,bitmapMissingFonts
					,useISO19005_1
					};

			Dispatch.invoke(dispatchObj, PDFCommandConst.EXPORTASFIXEDFORMAT,
					Dispatch.Method, args, new int[1]);

			Dispatch.call(dispatchObj, PDFCommandConst.CLOSE);

		} catch (ComException pdfOxyExcp) {
			PDFErrorPojo errorObj = PDFErrorGenerator.generateErrorMsg(
					pdfOxyExcp.getMessage(), terminate);
			if(errorObj.getErrorCode() == 1012){
				errorObj.setErrorDescription(PDFErrorMessages.PDF_CV_NO_MS_PUB);
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
