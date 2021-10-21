/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   16-11-2017 
 */

package com.zeonpad.pdfcovertor;

public abstract interface PDFCommandConst {

	public static final String MS_WORD = "Word.Application";
	public static final String DOCUMENTS = "Documents";
	
	public static final String EXPORTASFIXEDFORMAT = "ExportAsFixedFormat";
	public static final String VISIBLE = "Visible";
	
	public static final String MS_EXCEL = "Excel.Application";
	public static final String WORKBOOKS = "Workbooks";

	public static final String MS_PPT = "Powerpoint.Application";
	public static final String PRESENTATIONS = "Presentations";
	
	public static final String MS_PUBLISHER = "Publisher.Application";
	public static final String DOCUMENT = "Document";
	
	
	public static final String MS_OUTLOOK = "Outlook.Application";
	public static final String SESSION = "Session";
	public static final String OPENSHAREDITEM="OpenSharedItem";
	
	public static final String MS_VISIO = "Visio.Application";
	
	public static final String OPEN = "Open";
	public static final String SAVEAS = "SaveAs";
	public static final String CLOSE = "Close";
	public static final String QUIT = "Quit";
	public static final String DISPLAYALERTS="DisplayAlerts";
	public static final boolean DISPLAY_ALERTS=false;
	
	public static final String KILL_PPT="taskkill -IM  powerpnt.exe -f";
	
	public static final String KILL_WORD="taskkill -IM  winword.exe -f";
	
	public static final String KILL_EXCEL="taskkill -IM  excel.exe -f";
	
	public static final String KILL_PUB="taskkill -IM  mspub.exe -f";
	
	public static final String KILL_OUTLOOK="taskkill -IM  outlook.exe -f";
	
	public static final String KILL_VISIO="taskkill -IM  visio.exe -f";

}
