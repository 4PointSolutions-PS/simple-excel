package com._4point.utils.simpleexcel.api;

import java.util.Optional;

/**
 * Class represents a worksheet in Excel.  WOrksheets are retrieved using the Xlsx class.
 *
 */
public interface Worksheet extends Iterable<WorksheetRow>{

	Optional<WorksheetRow> getRow(int i);
	WorksheetRow createRow(int i);
}
