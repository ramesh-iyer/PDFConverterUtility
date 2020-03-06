package com.mm.converter.utility;

import java.util.ArrayList;
import java.util.List;

public class PdfConverter {

	// For testing only, remove it in production
	public static void main(String[] args) {

		List<String> files = new ArrayList<String>();
		files.add("C:\\Users\\USER\\Desktop\\resources\\resume1.gif");
		files.add("C:\\Users\\USER\\Desktop\\resources\\resume2.jpg");
		files.add("C:\\Users\\USER\\Desktop\\resources\\resume3.png");
		files.add("C:\\Users\\USER\\Desktop\\resources\\slerexe-company.TIF");
		files.add("C:\\Users\\USER\\Desktop\\resources\\resume_non-included_image.docx");
		files.add("C:\\Users\\USER\\Desktop\\resources\\resume_included_image.docx");
		files.add("C:\\Users\\USER\\Desktop\\resources\\sample.xls");
		files.add("C:\\Users\\USER\\Desktop\\resources\\resume4.pptx");

		String outFileDir = "C:\\Users\\USER\\Desktop\\resources";

		PdfConverter converter = new PdfConverter();

		for (String inFilePath : files) {

			try {
				converter.convertToPdf(inFilePath, outFileDir);
				Thread.sleep(10000);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}

	}
	// End
	
	

	/**
	 * Convert to PDF document
	 * 
	 * @param inFilePath
	 * @param outFileDir
	 */
	public void convertToPdf(final String inFilePath, final String outFileDir) {
		boolean result = false;

		switch (getFileExtension(inFilePath)) {

		case "jpeg":
		case "jpg":
		case "gif":
		case "tif":
		case "png":
			ImageToPDFConverter imageToPdfConverter = new ImageToPDFConverter();
			result = imageToPdfConverter.convertToPdf(inFilePath, outFileDir);
			break;

		case "doc":
		case "docx":
			WordToPDFConverter wordToPDFConverter = new WordToPDFConverter();
			result = wordToPDFConverter.convertToPdf(inFilePath, outFileDir);
			break;

		case "xls":
		case "xlsx":
			ExcelToPDFConverter excelToPDFConverter = new ExcelToPDFConverter();
			result = excelToPDFConverter.convertToPdf(inFilePath, outFileDir);
			break;

		case "ppt":
		case "pptx":
			try {
				PptToPDFConverter pptToPDFConverter = new PptToPDFConverter();
				result = pptToPDFConverter.convertToPdf(inFilePath, outFileDir, getFileExtension(inFilePath));
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		System.out.println(getFileExtension(inFilePath) + " to PDFConversion was? :: " + ((result) ? "successful" : "failed"));
	}

	/**
	 * Get file extension
	 * 
	 * @param fileName
	 * @return
	 */
	private static String getFileExtension(String fileName) {
		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
		else
			return "";
	}

}
