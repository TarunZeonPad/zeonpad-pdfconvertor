/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;

public abstract interface PDFNumConst {

	// Word to PDF format
	public static final int wordToFormatPDF=17;
	// Don't save the changes pending.
	public static final int wordDoNotSaveChanges = 0;
	// Excel to PDF format
	public static final int excelToFormatPDF = 0;
	// Ppt to PDF format
	public static final int pptToFormatPDF = 32;
	// Publisher to PDF format
	public static final int publisherToFormatPDF = 2;
	// Don't save the changes pending.
	public static final int publisherDoNotSaveChanges = 3;
	//The target for down-sampling of colored images. Measured in dots per inch. 
	public static final int colorDownsampleTarget=300;
	//The threshold at or above which an image is down-sampled to the ColorDownsample target level. 
	public static final int colorDownsampleThreshold=450; 
	//The target for down-sampling of one-bit images.
	public static final int oneBitDownsampleTarget=1200;
	//The threshold at or above which an image is down-sampled to the OneBitDownsample target level.
	public static final int OneBitDownsampleThreshold=1800;
	
	// PDF to word format
	public static final int pdfToFormatWord=12;
	
	public static final int wdOpenFormatAuto=0;
	
	public static final int wdFormatXMLDocument=12;
	
	public static final int wdFormatHTML=8;
	
	public static final int wdFormatEncodedText=7;
	
	

}
