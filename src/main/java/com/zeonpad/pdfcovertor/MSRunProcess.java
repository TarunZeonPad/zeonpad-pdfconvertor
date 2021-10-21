/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-01-2017 
 */
package com.zeonpad.pdfcovertor;

public class MSRunProcess {
	
	public static void runPrepare(String prepare) {
		try {
			System.out.println(prepare);
	    	Process process = Runtime.getRuntime().exec(prepare);
	    	System.out.println("After conversion...");
	    } catch (Exception e) {
	    	e.printStackTrace();
	    }
	}
	
}
