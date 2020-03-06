package com.mm.converter.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.Anchor;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chapter;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Section;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * Utility class for converting (.xls, .xlsx) into PDF format.
 * 
 * @author ramesh.iyer
 *
 */
public class ExcelToPDFConverter {

	/**
	 * @param args
	 * @throws Exception
	 */

	// Font
	private static Font catFont = new Font(Font.FontFamily.TIMES_ROMAN, 18, Font.BOLD);
	private static Font redFont = new Font(Font.FontFamily.TIMES_ROMAN, 12, Font.NORMAL, BaseColor.RED);
	private static Font subFont = new Font(Font.FontFamily.TIMES_ROMAN, 16, Font.BOLD);
	private static Font smallBold = new Font(Font.FontFamily.TIMES_ROMAN, 12, Font.BOLD);

	// Number of columns in the excel
	private static int numberOfColumns;

	public boolean convertToPdf(final String inFilePath, final String outFileDir) {
		boolean result = false;
		String outPutFile = outFileDir + "\\" + getFileNameWithoutExtension(inFilePath) + ".pdf";

		File xlsFile = new File(inFilePath);
		Workbook workbook;
		try {
			long start = System.currentTimeMillis();
			workbook = loadSpreadSheet(xlsFile);
			result = convertSpreadSheetToPdf(workbook, outPutFile);
			System.out.println("Converted to PDF file in :: " + (System.currentTimeMillis() - start) + " ms");
		} catch (FileNotFoundException e) {
			System.exit(1);
		} catch (Exception e) {
			e.printStackTrace();
		}

		return result;
	}

	private boolean convertSpreadSheetToPdf(Workbook workbook, final String outPutFile)
			throws IOException, DocumentException {
		boolean result = false;
		Document document = new Document();
		PdfWriter.getInstance(document, new FileOutputStream(outPutFile));
		document.open();
		addMetaData(document);
		addTitlePage(document);

		Anchor anchor = new Anchor("Chapter 1", catFont);
		anchor.setName("Chapter 1");

		// Second parameter is the number of the chapter
		Chapter catPart = new Chapter(new Paragraph(anchor), 1);

		Paragraph subPara = new Paragraph("Table", subFont);
		Section subCatPart = catPart.addSection(subPara);
		addEmptyLine(subPara, 5);

		Sheet sheet = workbook.getSheetAt(0);

		// Iterate through each rows from first sheet
		Iterator<Row> rowIterator = sheet.iterator();

		int temp = 0;
		boolean flag = true;
		PdfPTable table = null;

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			int cellNumber = 0;

			if (flag) {
				table = new PdfPTable(row.getLastCellNum());
				flag = false;
			}

			// For each row, iterate through each columns
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					if (temp == 0) {
						numberOfColumns = row.getLastCellNum();
						PdfPCell c1 = new PdfPCell(new Phrase(cell.getStringCellValue()));
						c1.setHorizontalAlignment(Element.ALIGN_CENTER);
						table.addCell(c1);
						table.setHeaderRows(1);

					} else {
						cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell, table);
					}
					cellNumber++;
					break;

				case Cell.CELL_TYPE_NUMERIC:
					cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell, table);
					cellNumber++;
					break;
				}
			}
			temp = 1;
			if (numberOfColumns != cellNumber) {
				for (int i = 0; i < (numberOfColumns - cellNumber); i++) {
					table.addCell(" ");
				}
			}
		}
		subCatPart.add(table);
		// Add all this to the document
		document.add(catPart);
		document.close();
		result = true;

		return result;
	}

	/**
	 * Checks EmptyCell and Adding Cell Content
	 * 
	 * @param cellNumber
	 * @param cell
	 * @param table
	 * @return
	 */
	private static int checkEmptyCellAndAddCellContentToPDFTable(int cellNumber, Cell cell, PdfPTable table) {
		if (cellNumber == cell.getColumnIndex()) {
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				table.addCell(Double.toString(cell.getNumericCellValue()));
			}
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				table.addCell(cell.getStringCellValue());
			}

		} else {
			while (cellNumber < cell.getColumnIndex()) {

				table.addCell(" ");
				cellNumber++;

			}
			if (cellNumber == cell.getColumnIndex()) {
				if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					table.addCell(Double.toString(cell.getNumericCellValue()));
				}
				if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					table.addCell(cell.getStringCellValue());
				}

			}
			cellNumber = cell.getColumnIndex();
		}

		return cellNumber;
	}

	/**
	 * Adding Metadata and other details of the document
	 * 
	 * @param document
	 */
	private static void addMetaData(Document document) {
		document.addTitle(" ");
		document.addSubject(" ");
		document.addKeywords(" ");
		document.addAuthor(" ");
		document.addCreator(" ");
	}

	/**
	 * Adding Title of the document
	 * 
	 * @param document
	 * @throws DocumentException
	 */
	private static void addTitlePage(Document document) throws DocumentException {
		Paragraph preface = new Paragraph();
		// add one empty line
		addEmptyLine(preface, 1);
		// Big header
		preface.add(new Paragraph(" ", catFont));

		addEmptyLine(preface, 1);
		// Report generated by:
		preface.add(new Paragraph(" ", smallBold));
		addEmptyLine(preface, 3);
		preface.add(new Paragraph(" ", smallBold));
		addEmptyLine(preface, 8);
		preface.add(new Paragraph(" ", smallBold));

		document.add(preface);
		// Start a new page
		document.newPage();
	}

	/**
	 * Adding empty line
	 * 
	 * @param paragraph
	 * @param number
	 */
	private static void addEmptyLine(Paragraph paragraph, int number) {
		for (int i = 0; i < number; i++) {
			paragraph.add(new Paragraph(" "));
		}
	}

	/**
	 * Loading spread sheet
	 * 
	 * @param xlsFile
	 * @return
	 * @throws Exception
	 */
	private static Workbook loadSpreadSheet(File xlsFile) throws Exception {
		Workbook workbook = null;

		String ext = getFileExtension(xlsFile.getName());
		if (ext.equalsIgnoreCase("xlsx")) {
			OPCPackage pkg = OPCPackage.open(xlsFile.getAbsolutePath());
			workbook = new XSSFWorkbook(pkg);
			pkg.close();
		} else if (ext.equalsIgnoreCase("xls")) {
			InputStream xlsFIS = new FileInputStream(xlsFile);
			workbook = new HSSFWorkbook(xlsFIS);
			xlsFIS.close();
		} else {
			throw new Exception("File extension is not regonised");
		}
		return workbook;
	}

	/**
	 * Get File extension
	 * 
	 * @param fileName
	 * @return
	 */
	private static String getFileExtension(String fileName) {
		String ext = "";
		int mid = fileName.lastIndexOf(".");
		ext = fileName.substring(mid + 1, fileName.length());
		return ext;
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
