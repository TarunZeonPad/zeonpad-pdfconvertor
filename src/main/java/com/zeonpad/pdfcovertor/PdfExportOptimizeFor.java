/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   17-11-2017 
 */

package com.zeonpad.pdfcovertor;

public abstract interface PdfExportOptimizeFor {
	//Export for screen, which is a lower quality and results in a smaller file size.
	public static final int PDF_EXPORT_OPTIMIZE_FOR_ONSCREEN=1;
	//Export for print, which is higher quailty and results in a larger file size.
	public static final int PDF_EXPORT_OPTIMIZE_FORPRINT=0;
}
