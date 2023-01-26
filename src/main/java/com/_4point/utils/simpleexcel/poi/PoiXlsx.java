package com._4point.utils.simpleexcel.poi;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Objects;
import java.util.Optional;

import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com._4point.utils.simpleexcel.api.Worksheet;
import com._4point.utils.simpleexcel.api.Xlsx;

public class PoiXlsx implements Xlsx {

	private final Workbook xlsxDocument;
	private final CellStyle boldStyle;
	private final CellStyle centeredBoldStyle;

	public PoiXlsx() {
		this.xlsxDocument = new XSSFWorkbook();
		this.boldStyle = this.xlsxDocument.createCellStyle();		
		this.centeredBoldStyle = this.xlsxDocument.createCellStyle();
		
		Font boldFont = this.xlsxDocument.createFont();
		boldFont.setBold(true);
		
		boldStyle.setFont(boldFont);
		
		centeredBoldStyle.setFont(boldFont);
        centeredBoldStyle.setAlignment(HorizontalAlignment.CENTER);
	}

	/**
	 * Private constructor - use Xlsx.read() to instantiate this class.
	 * 
	 * @param doc
	 */
	private PoiXlsx(Workbook doc) {
		super();
		this.xlsxDocument = doc;
		this.boldStyle = null;		
		this.centeredBoldStyle = null;
	}
	

	public static Xlsx read(Path xlsxFile) throws XlsxException {
		if (Files.notExists(xlsxFile)) {
			throw new XlsxException("Excel file '" + xlsxFile.toString() + "' does not exist!");
		}
		try {
			return internalRead(Files.newInputStream(xlsxFile));
		} catch (IOException ioe) {
			throw new XlsxException("Excel file '" + xlsxFile.toString() + "' could not be read. " + ioe.getMessage(), ioe);
		} catch (UnsupportedFileFormatException uffe) {
			throw new XlsxException("Excel file  '" + xlsxFile.toString() + "' is not supported file format.", uffe);
		}
	}

	/**
	 * Reads in an XLSX from an InputStream, creates an XLSX object.
	 * 
	 * @param xlsxStream
	 * @return
	 * @throws XlsxException
	 * @throws IOException
	 */
	public static Xlsx read(InputStream xlsxStream) throws XlsxException, UnsupportedFileFormatException, IOException {
		try {
			return internalRead(xlsxStream);
		} catch (UnsupportedFileFormatException uffe) {
			throw new XlsxException("The input file is not a valid XLSX Document.", uffe);
		}
	}

	public static Xlsx create() {
		return new PoiXlsx();

	}
	
	@Override
	public Optional<Worksheet> getSheet(String worksheetName) {
		Sheet sheet = this.xlsxDocument.getSheet(worksheetName);
		return sheet != null ? Optional.of(new PoiWorksheet(sheet)) : Optional.empty();
	}

	@Override
	public Worksheet createSheet(String worksheetName) {
		Sheet sheet = Objects.requireNonNull(this.xlsxDocument.createSheet(worksheetName), ()->"Couldn't create worksheet named '" + worksheetName + "' within the provided spreadsheet.");		
		return new PoiWorksheet(sheet);
	}

	// This method is called by the public read methods, but throws a larger selection of exceptions so that we can produce
	// a more meaningful error message on the XdpException.
	private static Xlsx internalRead(InputStream xlsxStream) throws IOException, UnsupportedFileFormatException, XlsxException {
		// Parse the input stream into a XLSX Document object
        Workbook doc = new XSSFWorkbook(xlsxStream);
		return new PoiXlsx(doc);
	}
	
	public void write(OutputStream xlsxStream) throws IOException {
		xlsxDocument.write(xlsxStream);
	}

	@Override
	public void save(Path filePath) throws IOException {
		try (final OutputStream os = Files.newOutputStream(filePath, StandardOpenOption.CREATE)) {
			xlsxDocument.write(os);
			xlsxDocument.close();
		}
	}
}
