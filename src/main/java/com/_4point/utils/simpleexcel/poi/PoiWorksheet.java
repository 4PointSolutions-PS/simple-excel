package com._4point.utils.simpleexcel.poi;

import java.util.Iterator;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com._4point.utils.simpleexcel.api.Worksheet;
import com._4point.utils.simpleexcel.api.WorksheetRow;

public class PoiWorksheet implements Worksheet {

	private Sheet sheet;
	
	public PoiWorksheet(Sheet sheet) {
		this.sheet = sheet;
	}

	@Override
	public Optional<WorksheetRow> getRow(int rowNum) {
		final Row row = sheet.getRow(rowNum);
		return row != null ? Optional.of(new PoiWorksheetRow(row)) : Optional.empty();
	}

	@Override
	public Iterator<WorksheetRow> iterator() {
		return new WorksheetRowIterator(sheet.iterator());
	}

	private static class WorksheetRowIterator implements Iterator<WorksheetRow> {

		private Iterator<Row> iterator;
		
		public WorksheetRowIterator(Iterator<Row> iterator) {
			this.iterator = iterator;
		}

		@Override
		public boolean hasNext() {
			return iterator.hasNext();
		}

		@Override
		public WorksheetRow next() {
			return new PoiWorksheetRow(iterator.next());
		}
		
	}

	@Override
	public WorksheetRow createRow(int i) {
		final Row row = sheet.createRow(i);
		return row != null ? new PoiWorksheetRow(row) : null;
	}
}
