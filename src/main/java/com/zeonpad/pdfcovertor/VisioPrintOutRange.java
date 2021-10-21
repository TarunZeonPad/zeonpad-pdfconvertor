/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   30-11-2017 
 */

package com.zeonpad.pdfcovertor;

public abstract interface VisioPrintOutRange {

	// Prints all foreground pages.
	public static final int VISIO_PRINT_ALL = 0;

	// Prints the active page.
	public static final int VISIO_PRINT_CURRENT_PAGE = 2;
	
	//Prints the current view area.
	public static final int VISIO_PRINT_CURRENT_VIEW = 4;

	//Prints pages between the FromPage value and the ToPage value.
	public static final int VISIO_PRINT_FROM_TO = 4;
	
	//Prints a selection.
	public static final int VISIO_PRINT_SELECTION = 3;

}
