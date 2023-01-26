package com._4point.utils.simpleexcel.api;

import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.util.Date;
import java.util.Optional;

/**
 * This class represents a row within a Worksheet.  WorksheetRows are retrieved from the Worksheet class.
 *
 * This class only has a couple of ways of retrieving data from columns.  It is expected that more may be
 * added over time.
 * 
 * The class also uses the older Java Date class because that is what Apache POI returns.
 * 
 * This class has some convenience functions so that new code does not have to deal with Date objects.
 *
 */
public interface WorksheetRow {

	Optional<String> getColumnValueAsString(int i);

	Optional<Date> getColumnValueAsDate(int i);

	WorksheetRow setColumnValueAsString(int i, String value);
	
	WorksheetRow setColumnValueAsDate(int i, Date value);
	
	default Optional<LocalDate> getColumnValueAsLocalDate(int i) {
		return getColumnValueAsDate(i).map(date->LocalDate.parse(new SimpleDateFormat("yyyy-MM-dd").format(date)));
	}
	
	default Optional<Instant> getColumnValueAsInstant(int i) {
		return getColumnValueAsDate(i).map(date->date.toInstant());
	}
	
	default WorksheetRow setColumnValueAsDate(int i, Instant value) {
		return setColumnValueAsDate(i, Date.from(value));
	}
}
