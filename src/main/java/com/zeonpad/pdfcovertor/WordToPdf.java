/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   16-11-2017 
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

public class WordToPdf {

	private ActiveXComponent pdfApp;
	private Dispatch pdfDispatch;

	private ScheduledExecutorService pdfScheduler = null;
	private ScheduledFuture<?> knightWatch = null;

	private boolean openAfterExport;
	private int pdfExportOptimizeFor;
	private int pdfExportRange;
	private int startPage;
	private int endPage;
	private int documentMarkups;
	private boolean includeDocProps;
	private boolean paramKeepIRM = true;
	private int createBookmarks;
	private boolean docStructureTags = true;
	private boolean bitmapMissingFonts = true;
	private boolean useISO19005_1 = false;
	private boolean terminate;
	// Default Set to 5 minute.
	private int conversionTimeOut = 5;
	// Default Password Set to ZeonPassword
	private String documentPassword = "ZeonPassword";

	/**
	 * @return the pdfExportOptimizeFor
	 */
	public int getPdfExportOptimizeFor() {
		return pdfExportOptimizeFor;
	}

	/**
	 * @param pdfExportOptimizeFor
	 *            the pdfExportOptimizeFor to set
	 */
	public void setPdfExportOptimizeFor(int pdfExportOptimizeFor) {

		if (pdfExportOptimizeFor == PdfExportOptimizeFor.PDF_EXPORT_OPTIMIZE_FOR_ONSCREEN
				|| pdfExportOptimizeFor == PdfExportOptimizeFor.PDF_EXPORT_OPTIMIZE_FORPRINT) {
			this.pdfExportOptimizeFor = pdfExportOptimizeFor;
		}

	}

	/**
	 * @return the pdfExportRange
	 */
	public int getPdfExportRange() {
		return pdfExportRange;
	}

	/**
	 * @param pdfExportRange
	 *            the pdfExportRange to set
	 */
	public void setPdfExportRange(int pdfExportRange) {
		if (pdfExportRange == PdfExportRange.PDF_EXPORT_ALL_DOCUMENT
				|| pdfExportRange == PdfExportRange.PDF_EXPORT_SELECTION
				|| pdfExportRange == PdfExportRange.PDF_EXPORT_CURRENT_PAGE
				|| pdfExportRange == PdfExportRange.PDF_EXPORT_FROM_TO) {
			this.pdfExportRange = pdfExportRange;
		}
	}

	/**
	 * @return the startPage
	 */
	public int getStartPage() {
		return startPage;
	}

	/**
	 * @param startPage
	 *            the startPage to set
	 */
	public void setStartPage(int startPage) {
		if (startPage > 0) {
			this.startPage = startPage;
		}

	}

	/**
	 * @return the endPage
	 */
	public int getEndPage() {
		return endPage;
	}

	/**
	 * @param endPage
	 *            the endPage to set
	 */
	public void setEndPage(int endPage) {
		if (endPage > 0) {
			this.endPage = endPage;
		}

	}

	/**
	 * @return the documentMarkups
	 */
	public int getDocumentMarkups() {
		return documentMarkups;
	}

	/**
	 * @param documentMarkups
	 *            the documentMarkups to set
	 */
	public void setDocumentMarkups(int documentMarkups) {
		if (documentMarkups == PDFMarkups.PDF_EXPORT_WITH_MARKUP
				|| documentMarkups == PDFMarkups.PDF_EXPORT_WITHOUT_MARKUP) {
			this.documentMarkups = documentMarkups;
		}

	}

	/**
	 * @return the paramKeepIRM
	 */
	public boolean isParamKeepIRM() {
		return paramKeepIRM;
	}

	/**
	 * @param paramKeepIRM
	 *            the paramKeepIRM to set
	 */
	public void setParamKeepIRM(boolean paramKeepIRM) {
		this.paramKeepIRM = paramKeepIRM;
	}

	/**
	 * @return the createBookmarks
	 */
	public int getCreateBookmarks() {
		return createBookmarks;
	}

	/**
	 * @param createBookmarks
	 *            the createBookmarks to set
	 */
	public void setCreateBookmarks(int createBookmarks) {
		if (createBookmarks == PdfCreateBookmark.PDF_CREATE_NO_BOOKMARK
				|| createBookmarks == PdfCreateBookmark.PDF_CREATE_HEADINGBOOKMARKS
				|| createBookmarks == PdfCreateBookmark.PDF_CREATE_WORDBOOKMARKS) {
			this.createBookmarks = createBookmarks;
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
	 * @return the docStructureTags
	 */
	public boolean isDocStructureTags() {
		return docStructureTags;
	}

	/**
	 * @param docStructureTags
	 *            the docStructureTags to set
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
	 * @param bitmapMissingFonts
	 *            the bitmapMissingFonts to set
	 */
	public void setBitmapMissingFonts(boolean bitmapMissingFonts) {
		this.bitmapMissingFonts = bitmapMissingFonts;
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

	/**
	 * @param documentPassword
	 *            the documentPassword to set
	 */
	public void setDocumentPassword(String documentPassword) {
		this.documentPassword = documentPassword;
	}

	public void convert(String source, String destination) throws PDFException {
		
		PDFileValidation.validateFile(destination);
		ComThread.InitSTA();
		try {

			if (pdfExportRange == PdfExportRange.PDF_EXPORT_FROM_TO
					&& startPage == 0) {
				startPage = 1;
			}

			if (pdfExportRange == PdfExportRange.PDF_EXPORT_FROM_TO
					&& endPage == 0) {
				throw new PDFException(
						PDFErrorMessages.PDF_CV_END_VAL_NDF_CODE,
						PDFErrorMessages.PDF_CV_END_VAL_NDF);
			}

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
			pdfApp.setProperty(PDFCommandConst.DISPLAYALERTS,
					PDFCommandConst.DISPLAY_ALERTS);
			pdfDispatch = pdfApp.getProperty(PDFCommandConst.DOCUMENTS)
					.toDispatch();

			Dispatch dispatchObj = Dispatch.call(pdfDispatch,
					PDFCommandConst.OPEN, source, false, true, false,
					documentPassword).toDispatch();

			// Dispatch.call(pdfDispatch, "Unprotect", "password");

			Object[] args = new Object[] { destination,
					new Integer(PDFNumConst.wordToFormatPDF),
					new Boolean(openAfterExport),
					new Integer(pdfExportOptimizeFor),
					new Integer(pdfExportRange), new Integer(startPage),
					new Integer(endPage), new Integer(documentMarkups),
					new Boolean(includeDocProps), new Boolean(paramKeepIRM),
					new Integer(createBookmarks),
					new Boolean(docStructureTags),
					new Boolean(bitmapMissingFonts), new Boolean(useISO19005_1) };

			Dispatch.invoke(dispatchObj, PDFCommandConst.EXPORTASFIXEDFORMAT,
					Dispatch.Method, args, new int[1]);

			Dispatch.call(dispatchObj, PDFCommandConst.CLOSE, false);

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
