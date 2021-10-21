package com.zeonpad.wordcovertor;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.zeonpad.pdfcovertor.PDFCommandConst;

public class POC {
	
	// Declare word object
		private ActiveXComponent objWord;

		// Declare Word Properties
		private Dispatch custDocprops;

		private Dispatch builtInDocProps;

		// the doucments object is important in any real app but this demo doesn't
		// use it
		// private Dispatch documents;

		private Dispatch document;

		private Dispatch wordObject;

	
	public static void main(String[] args) {
		
		POC obj = new POC();
		//obj.open("C:\\Users\\tarun.kumar\\Desktop\\RESTful Java Patterns and Best Practices.pdf");
		wordToHtml("C:\\Users\\tarun.kumar\\Desktop\\RESTful Java Patterns and Best Practices.pdf", "C:\\Users\\tarun.kumar\\Desktop\\A.docx");
		
	}

	
	public static final int WORD_HTML = 8;
	public static final int WORD_TXT = 7;
	public static final int EXCEL_HTML = 44;
	public static boolean wordToHtml(String fileDoc, String fileHtml) {
		ActiveXComponent app = new ActiveXComponent("Word.Application"); // Start word
		try {
			// Set the word invisible
			app.setProperty("Visible", new Variant(false));
			app.setProperty("ScreenUpdating",
					 new Variant(false));
			app.setProperty("DisplayAlerts", new Variant(0));
			System.out.println("OPening the file....1");
			// Get a documents object
			Dispatch docs = (Dispatch) app.getProperty("Documents")
					.toDispatch();
			System.out.println("OPening the file....2");
			// Open the file
			Dispatch doc = Dispatch.invoke(
					docs,
					"Open",
					Dispatch.Method,
					new Object[] { fileDoc, new Variant(false),
							new Variant(false),new Variant(false),"","",new Variant(false),"","",0 }, new int[1]).toDispatch();
			
			System.out.println("OPening the file....3");
			// Save the new file
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] {
					fileHtml, new Variant(12),new Variant(false),"",new Variant(true),"",new Variant(false),new Variant(false),new Variant(false),new Variant(false),new Variant(false) }, new int[1]);
			Variant f = new Variant(false);
			Dispatch.call(doc, "Close", f);
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			app.invoke("Quit", new Variant[] {});
		}
	}
	
	public void open(String filename) {
		// Instantiate objWord
		objWord = new ActiveXComponent("Word.Application");

		// Assign a local word object
		wordObject = objWord.getObject();

		// Create a Dispatch Parameter to hide the document that is opened
		Dispatch.put(wordObject, "Visible", new Variant(false));

		// Instantiate the Documents Property
		Dispatch documents = objWord.getProperty("Documents").toDispatch();

		// Open a word document, Current Active Document
		document = Dispatch.call(documents, "Open", filename).toDispatch();
	}
}
