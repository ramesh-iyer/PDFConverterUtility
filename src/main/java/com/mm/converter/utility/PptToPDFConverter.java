package com.mm.converter.utility;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * Utility class for converting (.ppt, .pptx) into PDF format.
 * 
 * @author ramesh.iyer
 *
 */

public class PptToPDFConverter {

	/**
	 * Method for converting (.ppt, .pptx) into PDF format.
	 * 
	 * @param inFilePath
	 * @param outFileDir
	 * @param fileType
	 * @return
	 * @throws Exception
	 */
	public boolean convertToPdf(String inFilePath, String outFileDir, String fileType) throws Exception {
		boolean result = false;
		String outputFileName = getFileNameWithoutExtension(inFilePath) + ".pdf";

		FileInputStream inputStream = new FileInputStream(inFilePath);
		double zoom = 2;
		AffineTransform at = new AffineTransform();
		at.setToScale(zoom, zoom);
		long start = System.currentTimeMillis();
		Document pdfDocument = new Document();
		PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument,
				new FileOutputStream(outFileDir + "//" + outputFileName));
		PdfPTable table = new PdfPTable(1);
		pdfWriter.open();
		pdfDocument.open();
		Dimension pgsize = null;
		Image slideImage = null;
		BufferedImage img = null;
		if (fileType.equalsIgnoreCase("ppt")) {
			SlideShow ppt = new SlideShow(inputStream);
			pgsize = ppt.getPageSize();
			Slide slide[] = ppt.getSlides();
			pdfDocument.setPageSize(new Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
			pdfWriter.open();
			pdfDocument.open();
			for (int i = 0; i < slide.length; i++) {
				img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom),
						BufferedImage.TYPE_INT_RGB);
				Graphics2D graphics = img.createGraphics();
				graphics.setTransform(at);

				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
				slide[i].draw(graphics);
				graphics.getPaint();
				slideImage = Image.getInstance(img, null);
				table.addCell(new PdfPCell(slideImage, true));
			}

		}
		if (fileType.equalsIgnoreCase("pptx")) {
			XMLSlideShow ppt = new XMLSlideShow(inputStream);
			pgsize = ppt.getPageSize();
			XSLFSlide slide[] = ppt.getSlides();
			pdfDocument.setPageSize(new Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
			pdfWriter.open();
			pdfDocument.open();
			for (int i = 0; i < slide.length; i++) {
				img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom),
						BufferedImage.TYPE_INT_RGB);
				Graphics2D graphics = img.createGraphics();
				graphics.setTransform(at);

				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
				slide[i].draw(graphics);
				graphics.getPaint();
				slideImage = Image.getInstance(img, null);
				table.addCell(new PdfPCell(slideImage, true));
			}
		}
		pdfDocument.add(table);
		pdfDocument.close();
		pdfWriter.close();
		result = true;
		System.out.println("Converted to PDF file in :: " + (System.currentTimeMillis() - start) + " ms");
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
