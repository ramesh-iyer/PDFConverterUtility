package com.mm.converter.utility;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.nio.file.Path;
import java.nio.file.Paths;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * Utility class for converting (.jpeg, .gif, .tiff, .png) into PDF format.
 * 
 * @author ramesh.iyer
 *
 */

public class ImageToPDFConverter {

	/**
	 * Method for converting (.jpeg, .gif, .tiff, .png) into PDF format.
	 * 
	 * @param inFilePath
	 * @param outFileDir
	 * @return
	 */
	public boolean convertToPdf(final String inFilePath, final String outFileDir) {
		boolean result = false;
		try {

			File root = new File(outFileDir);
			String outputFileName = getFileNameWithoutExtension(inFilePath) + ".pdf";

			String inFileLocation = new File(inFilePath).getParent().toString();

			long start = System.currentTimeMillis();
			Document document = new Document();

			PdfWriter.getInstance(document, new FileOutputStream(new File(root, outputFileName)));

			document.open();

			document.newPage();
			Image image = Image.getInstance(new File(inFileLocation, getFileNameOnly(inFilePath)).getAbsolutePath());
			image.setAbsolutePosition(0, 0);
			image.setBorderWidth(0);
			image.scaleAbsoluteHeight(PageSize.A4.getHeight());
			image.scaleAbsoluteWidth(PageSize.A4.getWidth());
			document.add(image);
			result = true;
			System.out.println("Converted to a PDF file in :: " + (System.currentTimeMillis() - start) + " ms");
			document.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (DocumentException e) {
			e.printStackTrace();
		} catch (MalformedURLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return result;
	}

	private static String getFileNameOnly(String fileName) {
		Path path = Paths.get(fileName);
		return path.getFileName().toString();
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
