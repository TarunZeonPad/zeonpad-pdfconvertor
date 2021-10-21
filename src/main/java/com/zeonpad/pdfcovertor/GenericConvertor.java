/**
 * @author  Tarun Kumar 
 * @version 1.0
 * @since   12-06-2017 
 */
package com.zeonpad.pdfcovertor;

public class GenericConvertor {

	public void convert(String source, String destination) throws PDFException {

		PDFileValidation.validateInputFile(source);

		String fileExtension = getFileExtension(source);
		if (fileExtension.startsWith(".do") || fileExtension.startsWith(".tx")
				|| fileExtension.startsWith(".xm")
				|| fileExtension.startsWith(".ht")
				|| fileExtension.startsWith(".odt")
				|| fileExtension.startsWith(".mh")
				|| fileExtension.startsWith(".rt")) {
			WordToPdf wordToPdf = new WordToPdf();
			wordToPdf.convert(source, destination);
		} else if (fileExtension.startsWith(".xl")
				|| fileExtension.startsWith(".cs")
				|| fileExtension.startsWith(".ods")) {
			ExcelToPdf excToPdf = new ExcelToPdf();
			excToPdf.convert(source, destination);
		} else if (fileExtension.startsWith(".pp")
				|| fileExtension.startsWith(".po")
				|| fileExtension.startsWith(".odp")) {
			PptToPdf pptToPdf = new PptToPdf();
			pptToPdf.convert(source, destination);
		} else if (fileExtension.startsWith(".ms")) {
			OutlookToPdf outlookToPdf = new OutlookToPdf();
			outlookToPdf.convert(source, destination);
		} else if (fileExtension.startsWith(".pu")) {
			PublisherToPdf publisherToPdf = new PublisherToPdf();
			publisherToPdf.convert(source, destination);
		} else if (fileExtension.startsWith(".bmp")
				|| fileExtension.startsWith(".gif")
				|| fileExtension.startsWith(".jfif")
				|| fileExtension.startsWith(".jp")
				|| fileExtension.startsWith(".png")
				|| fileExtension.startsWith(".tif")) {
			ImageToPDF imageToPdf = new ImageToPDF();
			imageToPdf.setConversionTimeOut(1);
			imageToPdf
					.convert(
							source,
							destination);
		} else if (fileExtension.startsWith(".vd")
				|| fileExtension.startsWith(".vs")
				|| fileExtension.startsWith(".vt")
				|| fileExtension.startsWith(".sv")
				) {
			VisioToPdf visioToPdf = new VisioToPdf();
			visioToPdf.convert(source, destination);

		}else if(fileExtension.startsWith(".dw")
				|| fileExtension.startsWith(".dx")){
			
			AutocadToPdf autocadToPDf = new AutocadToPdf();
			autocadToPDf.convert(source, destination);
		}

	}

	private String getFileExtension(String fileName) {

		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0) {
			return fileName.substring(fileName.lastIndexOf(".")).toLowerCase();
		}

		return "";

	}

}
