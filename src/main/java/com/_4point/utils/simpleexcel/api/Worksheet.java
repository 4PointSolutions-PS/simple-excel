package com._4point.utils.simpleexcel.api;

import java.util.Optional;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Class represents a worksheet in Excel.  WOrksheets are retrieved using the Xlsx class.
 *
 */
public interface Worksheet extends Iterable<WorksheetRow>{

	Optional<WorksheetRow> getRow(int i);
	WorksheetRow createRow(int i);

	public default Stream<WorksheetRow> stream() {
		// This is a simplistic approach and could be improved using sheet.getLastRowNum() to implement a better spliterator, 
		// but it works for now.
		return StreamSupport.stream(this.spliterator(), false);
	}
}
