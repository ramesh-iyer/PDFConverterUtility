package com.mm.converter.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Utility class for converting (.doc, .docx) into PDF format.
 * 
 * @author ramesh.iyer
 *
 */
public class WordToPDFConverter {

	/**
	 * Method for converting (.doc, .docx) into PDF format.
	 * 
	 * @param inFilePath
	 * @param outFileDir
	 * @return
	 */
	public boolean convertToPdf(final String inFilePath, final String outFileDir) {
		boolean result = false;

		String outputFileName = getFileNameWithoutExtension(inFilePath) + ".pdf";
		try {
			InputStream is = new FileInputStream(new File(inFilePath));
			OutputStream out = new FileOutputStream(new File(outFileDir + "//" + outputFileName));

			long start = System.currentTimeMillis();
			// 1) Load DOCX into XWPFDocument
			XWPFDocument document = new XWPFDocument(is);

			// 2) Prepare Pdf options
			PdfOptions options = PdfOptions.create();

			// 3) Convert XWPFDocument to Pdf
			PdfConverter.getInstance().convert(document, out, options);
			result = true;

			System.out.println("Converted to PDF file in :: " + (System.currentTimeMillis() - start) + " ms");
		} catch (Throwable e) {
			e.printStackTrace();
		}
		return result;
	}

	/**
	 * Returns filename only
	 * 
	 * @param filePath
	 * @return
	 */
	private static String getFileNameWithoutExtension(String filePath) {
		String fileName = "";
		try {
			File file = new File(filePath);
			if (file != null && file.exists()) {
				String name = file.getName();
				fileName = name.replaceFirst("[.][^.]+$", "");
			}
		} catch (Exception e) {
			e.printStackTrace();
			fileName = "";
		}

		return fileName;

	}

}