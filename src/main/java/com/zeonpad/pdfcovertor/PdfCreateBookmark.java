/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;

public abstract interface PdfCreateBookmark {

	//Do not create bookmarks in the exported pdf document.
	public static final int PDF_CREATE_NO_BOOKMARK=0;
	
	//Create a bookmark in the exported pdf document for each Microsoft Office Word heading, which includes only headings within the main document and text boxes not within headers, footers, endnotes, footnotes, or comments.
	public static final int PDF_CREATE_HEADINGBOOKMARKS=1;
	
	//Create a bookmark in the exported pdf document for each Word bookmark, which includes all bookmarks except those contained within headers and footers.
	public static final int PDF_CREATE_WORDBOOKMARKS=2;
}
