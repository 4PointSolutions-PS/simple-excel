package com._4point.utils.simpleexcel.poi;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Date;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com._4point.utils.simpleexcel.api.Worksheet;
import com._4point.utils.simpleexcel.api.WorksheetRow;
import com._4point.utils.simpleexcel.api.Xlsx;
import com.google.common.jimfs.Configuration;
import com.google.common.jimfs.Jimfs;

class PoiXlsxTest {
	private static final boolean SAVE_RESULTS = true;
	
	private static final Path RESOURCES_DIR = Path.of("src", "test", "resources");
	private static final Path ACTUAL_RESULTS_DIR = RESOURCES_DIR.resolve("actualResults");
	
	private static final Path SAMPLE_EXCEL_FILE = RESOURCES_DIR.resolve("sample.xlsx");
	private static final String MAIN_WORKSHEET_NAME = "Forms Information";
	private static final String CHANNELS_WORKSHEET_NAME = "Effective_Expiry_Dates";

	@BeforeAll
	static void setUpBeforeClass() throws Exception {
	}

	@BeforeEach
	void setUp() throws Exception {
	}

	@Test
	void testCreateNewAndSave() throws Exception {
		final String targetFilename = "testCreateNewAndSave.xlsx";
		final String worksheetName = "Test Sheet";
		final PoiXlsx poiXlsx = new PoiXlsx();
		final Worksheet sheet = poiXlsx.createSheet(worksheetName);
		assertNotNull(sheet);
		
		final int rowIndex = 0;
		final int stringCellIndex = 0;
		final int dateCellIndex = 1;
		final int rewriteStringCellIndex = 2;
		final int rewriteDateCellIndex = 3;
		assertTrue(sheet.getRow(stringCellIndex).isEmpty());	// Newly created sheet should have no rows.

		final WorksheetRow row = sheet.createRow(rowIndex);
		assertNotNull(row);

		// Test the creation of a new cell
		final String stringCellValue = "Test Cell";
		final Date dateCellValue = new Date();
		row.setColumnValueAsString(stringCellIndex, stringCellValue);
		row.setColumnValueAsDate(dateCellIndex, dateCellValue);
		
		// Test the creation and then overwriting of a new cell
		final String initialStringCellValue = "Test Cell Intial";
		final Date initialDateCellValue = dateCellValue;
		row.setColumnValueAsString(rewriteStringCellIndex, initialStringCellValue);
		row.setColumnValueAsDate(rewriteDateCellIndex, initialDateCellValue);

		final String finalStringCellValue = "Test Cell Final";
		final Date finalDateCellValue = new Date();
		row.setColumnValueAsString(rewriteStringCellIndex, finalStringCellValue);
		row.setColumnValueAsDate(rewriteDateCellIndex, finalDateCellValue);

		final Path targetPath = Jimfs.newFileSystem(Configuration.unix()).getPath(targetFilename);
		poiXlsx.save(targetPath);
		
		assertTrue(Files.exists(targetPath));

		if (SAVE_RESULTS) {
			Files.copy(targetPath, ACTUAL_RESULTS_DIR.resolve(targetFilename), StandardCopyOption.REPLACE_EXISTING);
		}
		
		// Now that we've saved it, make sure we can read it back in
		final Xlsx createdXlsx = PoiXlsx.read(targetPath);
		final Worksheet createdSheet = createdXlsx.getSheet(worksheetName).get();
		final WorksheetRow createdRow = createdSheet.getRow(rowIndex).get();
		String createdValue1 = createdRow.getColumnValueAsString(stringCellIndex).get();
		assertEquals(stringCellValue, createdValue1);
		final Date createdValue2 = createdRow.getColumnValueAsDate(dateCellIndex).get();
		assertEquals(dateCellValue, createdValue2);

		String rewrittenValue1 = createdRow.getColumnValueAsString(rewriteStringCellIndex).get();
		assertEquals(finalStringCellValue, rewrittenValue1);
		final Date rewrittenValue2 = createdRow.getColumnValueAsDate(rewriteDateCellIndex).get();
		assertEquals(finalDateCellValue, rewrittenValue2);
}

	@Test
	void testLoadAndRead() throws Exception {
		final Xlsx poiXlsx = PoiXlsx.read(SAMPLE_EXCEL_FILE);
		
		final Worksheet sheet = poiXlsx.getSheet(MAIN_WORKSHEET_NAME).get();

		final WorksheetRow row = sheet.getRow(0).get();
		
		String value = row.getColumnValueAsString(0).get();
		assertEquals("FormId", value);
		
		// Test Iterator
		int count = 0;
		for (WorksheetRow r : sheet) {
			count++;
		}
		assertEquals(4, count);
		
		final Worksheet channelsSheet = poiXlsx.getSheet(CHANNELS_WORKSHEET_NAME).get();
		final WorksheetRow channelRow = channelsSheet.getRow(3).get();
		Date value1 = channelRow.getColumnValueAsDate(5).get();
		assertNotNull(value1);
		
	}

}
