package com._4point.utils.simpleexcel.poi;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;

import com._4point.utils.simpleexcel.api.WorksheetRow;

public class PoiWorksheetRow implements WorksheetRow {

	private static final SimpleDateFormat EXCEL_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

	private final Row row;
	
	public PoiWorksheetRow(Row row) {
		this.row = row;
	}

	@Override
	public Optional<String> getColumnValueAsString(int i) {
		String retVal = "";
		final Cell cell = row.getCell(i);
		if (cell == null) {
			return Optional.empty();
		}
		final CellType cellType = cell.getCellType();
		if (cellType == CellType.STRING) {
			retVal = cell.getStringCellValue();
		} else if (cellType == CellType.NUMERIC) {
			retVal = convertDoubleToString(cell.getNumericCellValue());
		} else if (cellType == CellType.BLANK) {
			return Optional.empty();
		} else {
			throw new IllegalStateException("Invalid Cell Type in speadsheet - column (" + i + ") - Cell Type (" + cellType.toString() + ").  Expected String.");
		}
		return Optional.of(retVal);
	}

	private String convertDoubleToString(double numericCellValue) {
		String retVal;
		if ((numericCellValue == Math.floor(numericCellValue)) && !Double.isInfinite(numericCellValue) &&
			numericCellValue < Long.MAX_VALUE && numericCellValue > Long.MIN_VALUE) {
				retVal = String.valueOf((long)numericCellValue);
		} else {
			retVal = String.valueOf(numericCellValue);
		}
		return retVal;
	}

	@Override
	public Optional<Date> getColumnValueAsDate(int i) {
		Date retVal = null;
		final Cell cell = row.getCell(i);
		if (cell == null) {
			return Optional.empty();
		}
		final CellType cellType = cell.getCellType();
		
		if (cellType == CellType.NUMERIC) {
//			System.out.println("Cell Style = '" + cell.getCellStyle().getDataFormatString() + "'.");
			retVal = cell.getDateCellValue();
		} else if (cellType == CellType.STRING) {
			final String stringCellValue = cell.getStringCellValue();
			try {
				retVal = EXCEL_DATE_FORMAT.parse(stringCellValue);
			} catch (ParseException e) {
				throw new IllegalStateException("Unable to parse date format in column (" + i + ").  Value is (" + stringCellValue + ").  Expected format to be '" + EXCEL_DATE_FORMAT + "' instead.");
			}
		} else if (cellType == CellType.BLANK) {
			return Optional.empty();
		} else {
			throw new IllegalStateException("Invalid Cell Type in speadsheet - column (" + i + ") - Cell Type (" + cellType.toString() + "). Expected Date.s");
		}
			
		return Optional.ofNullable(retVal);
	}

	@Override
	public PoiWorksheetRow setColumnValueAsString(int i, String value) {
		Cell cell = row.getCell(i);
		if (cell == null) {
			cell = row.createCell(i, CellType.STRING);
		}
		cell.setCellValue(value);
		return this;
	}

	@Override
	public WorksheetRow setColumnValueAsDate(int i, Date value) {
		Cell cell = row.getCell(i);
		if (cell == null) {
			cell = row.createCell(i, CellType.NUMERIC);
			DataFormat foo = row.getSheet().getWorkbook().createDataFormat();
			CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, foo.getFormat("yyyy\\-m\\-d\\ hh:mm"));
		}
		cell.setCellValue(value);
		return this;
	}	

}
