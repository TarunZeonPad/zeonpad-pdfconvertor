/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   17-11-2017 
 */

package com.zeonpad.pdfcovertor;

public abstract interface PdfExportRange {

	//Exports the entire document to pdf.
	public static final int PDF_EXPORT_ALL_DOCUMENT=0;
	
	//Exports the contents of the current selection.
	public static final int PDF_EXPORT_SELECTION=1;
	
	//Exports the current page.
	public static final int PDF_EXPORT_CURRENT_PAGE=2;
	
	//Exports the contents of a range using the starting and ending positions.
	public static final int PDF_EXPORT_FROM_TO=3;
}
