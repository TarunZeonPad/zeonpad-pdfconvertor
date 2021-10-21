/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   16-11-2017 
 */

package com.zeonpad.pdfcovertor;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

public class PDFConvertor {

	Dispatch pdfOxyDispatch;
	ActiveXComponent pdfOxyApp;

	public PDFConvertor() {

	}

	
	public void releaseResource(){
		 if (pdfOxyApp != null){
			 pdfOxyApp.invoke(PDFCommandConst.QUIT, PDFNumConst.wordDoNotSaveChanges);
		}
	}
}
