package com._4point.utils.simpleexcel.poi;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.Optional;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import com._4point.utils.simpleexcel.api.WorksheetRow;

class PoiWorksheetRowTest {

	@BeforeAll
	static void setUpBeforeClass() throws Exception {
	}

	@BeforeEach
	void setUp() throws Exception {
	}

	@Test
	void testGetColumnValueAsString_Integer() {
		int testValue = 23;
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockStringRow(testValue));
		final Optional<String> result = underTest.getColumnValueAsString(0);
		assertTrue(result.isPresent());
		assertEquals("23", result.get());
	}

	@Test
	void testGetColumnValueAsString_Fractional() {
		double testValue = 23.23;
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockStringRow(testValue));
		final Optional<String> result = underTest.getColumnValueAsString(0);
		assertTrue(result.isPresent());
		assertEquals("23.23", result.get());
	}
	
	@Test
	void testGetColumnValueAsString_String() {
		String testValue = "foo123";
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockStringRow(testValue));
		final Optional<String> result = underTest.getColumnValueAsString(0);
		assertTrue(result.isPresent());
		assertEquals(testValue, result.get());
	}
	
	@Test
	void testGetColumnValueAsString_Empty() {
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockStringRow());
		final Optional<String> result = underTest.getColumnValueAsString(0);
		assertTrue(result.isEmpty());
	}
	
	@Test
	void testGetColumnValueAsString_InvalidBoolean() throws Exception {
		final boolean testDateTime = false;
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow(testDateTime));
		IllegalStateException ex = assertThrows(IllegalStateException.class, ()->underTest.getColumnValueAsString(0));
		
		String msg = ex.getMessage();
		assertTrue(msg.contains("Invalid Cell Type"), "Missing 'Invalid Cell Type' in '" + msg + "'.");
		assertTrue(msg.contains("column (0)"), "Missing 'column (0)' in '" + msg + "'.");
		assertTrue(msg.contains("BOOLEAN"), "Missing 'BOOLEAN' in '" + msg + "'.");
	}

	@Test
	void testGetColumnValueAsDate_Date() {
		final Date testDate = Date.from(fromISODateTime("2011-12-03T10:15:30+01:00")); 
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow(testDate));
		final Optional<Date> result = underTest.getColumnValueAsDate(0);
		assertTrue(result.isPresent());
		assertEquals(testDate, result.get());
	}
	
	@Test
	void testGetColumnValueAsDate_String() throws Exception {
		final String testDateTime = "2011-12-03 10:15:30";
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow(testDateTime));
		final Optional<Date> result = underTest.getColumnValueAsDate(0);
		assertTrue(result.isPresent());
		assertEquals((new SimpleDateFormat("yyyy-MM-dd HH:mm:ss")).parse(testDateTime), result.get());
	}
	
	@Test
	void testGetColumnValueAsDate_InvalidString() throws Exception {
		final String testDateTime = "9:15:30 2011-12-03";
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow(testDateTime));
		IllegalStateException ex = assertThrows(IllegalStateException.class, ()->underTest.getColumnValueAsDate(0));
		
		String msg = ex.getMessage();
		assertTrue(msg.contains("Unable to parse date format"), "Missing 'Unable to parse date format' in '" + msg + "'.");
		assertTrue(msg.contains("column (0)"), "Missing 'column (0)' in '" + msg + "'.");
		assertTrue(msg.contains(testDateTime), "Missing test date '" + testDateTime + "' in '" + msg + "'.");
	}
	
	@Test
	void testGetColumnValueAsDate_InvalidBoolean() throws Exception {
		final boolean testDateTime = false;
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow(testDateTime));
		IllegalStateException ex = assertThrows(IllegalStateException.class, ()->underTest.getColumnValueAsDate(0));
		
		String msg = ex.getMessage();
		assertTrue(msg.contains("Invalid Cell Type"), "Missing 'Invalid Cell Type' in '" + msg + "'.");
		assertTrue(msg.contains("column (0)"), "Missing 'column (0)' in '" + msg + "'.");
		assertTrue(msg.contains("BOOLEAN"), "Missing 'BOOLEAN' in '" + msg + "'.");
	}
	
	@Test
	void testGetColumnValueAsDate_Empty() throws Exception {
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow());
		final Optional<Date> result = underTest.getColumnValueAsDate(0);
		assertTrue(result.isEmpty());
	}

	// This test requires more mocking that I am prepared to do at this time, so it is disabled.
	@Disabled
	void testSetColumnValueAsDate_Empty() throws Exception {
		final Date testDate = Date.from(fromISODateTime("1975-05-19T19:21:55+03:00")); 
		final PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow());
		assertTrue(underTest.getColumnValueAsDate(0).isEmpty());
		final WorksheetRow underTest2 = underTest.setColumnValueAsDate(0, testDate);
		final Optional<Date> result = underTest2.getColumnValueAsDate(0);
		assertTrue(result.isPresent());
		assertEquals(testDate, result.get());
	}
	
	@Test
	void testSetColumnValueAsString_Empty() {
		String testValue = "bar4567890";
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockStringRow());
		assertTrue(underTest.getColumnValueAsString(0).isEmpty());
		PoiWorksheetRow underTest2 = underTest.setColumnValueAsString(0, testValue);
		final Optional<String> result = underTest2.getColumnValueAsString(0);
		assertTrue(result.isPresent());
		assertEquals(testValue, result.get());
	}
	
	// This test requires more mocking that I am prepared to do at this time, so it is disabled.
	@Disabled
	void testSetColumnValueAsDate_NotEmpty() throws Exception {
		final Date initialDate = Date.from(fromISODateTime("1978-07-14T23:22:55+03:00")); 
		final Date testDate = Date.from(fromISODateTime("1975-05-19T19:21:55+03:00")); 
		final PoiWorksheetRow underTest = new PoiWorksheetRow(createMockDateRow(initialDate));
		Optional<Date> initialResult = underTest.getColumnValueAsDate(0);
		assertTrue(initialResult.isPresent());
		assertEquals(initialDate, initialResult.get());
		final WorksheetRow underTest2 = underTest.setColumnValueAsDate(0, testDate);
		final Optional<Date> result = underTest2.getColumnValueAsDate(0);
		assertTrue(result.isPresent());
		assertEquals(testDate, result.get());
	}
	
	@Test
	void testSetColumnValueAsString_NotEmpty() {
		String initialValue = "foobaz";
		String testValue = "bar4567890";
		PoiWorksheetRow underTest = new PoiWorksheetRow(createMockStringRow(initialValue));
		Optional<String> initialResult = underTest.getColumnValueAsString(0);
		assertTrue(initialResult.isPresent());
		assertEquals(initialValue, initialResult.get());
		PoiWorksheetRow underTest2 = underTest.setColumnValueAsString(0, testValue);
		final Optional<String> result = underTest2.getColumnValueAsString(0);
		assertTrue(result.isPresent());
		assertEquals(testValue, result.get());
	}
	
	
	private Instant fromISODateTime(String dateTime) {
		return Instant.from(DateTimeFormatter.ISO_OFFSET_DATE_TIME.parse(dateTime));
	}

	private Row createMockStringRow(int testValue) {
		return new MockRow() {
			private Cell cell = new MockCell() {
				private double value = (double) testValue;
				@Override
				public CellType getCellType() {
					return CellType.NUMERIC;
				}
				@Override
				public double getNumericCellValue() {
					return value;
				}
				@Override
				public void setCellValue(double value) {
					this.value = value;
				}
			};
			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
		};
	}

	private Row createMockStringRow(double testValue) {
		return new MockRow() {
			private Cell cell = new MockCell() {
				private double value = testValue;
				@Override
				public CellType getCellType() {
					return CellType.NUMERIC;
				}
				@Override
				public double getNumericCellValue() {
					return value;
				}
				@Override
				public void setCellValue(double value) {
					this.value = value;
				}
			};
			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
		};
	}

	private Row createMockStringRow(String testValue) {
		return new MockRow() {
			private Cell cell = new MockCell() {
				private String value = testValue;
				@Override
				public CellType getCellType() {
					return CellType.STRING;
				}
				@Override
				public String getStringCellValue() {
					return value;
				}
				@Override
				public void setCellValue(String value) {
					this.value = value;
				}
			};
			
			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
		};
	}

	private Row createMockStringRow() {
		return new MockRow() {
			private Cell cell = null;
			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
			@Override
			public Cell createCell(int column, CellType type) {
				if (type == CellType.STRING) {
					return cell = new MockCell() {
						private String value = null;
						@Override
						public CellType getCellType() {
							return CellType.STRING;
						}
						@Override
						public String getStringCellValue() {
							return value;
						}
						@Override
						public void setCellValue(String value) {
							this.value = value;
						}
					};
				} else {
					throw new IllegalArgumentException("Invalid CellType '" + type.toString() + "'.");
				}
			}
		};
	}

	private Row createMockDateRow(Date testValue) {
		return new MockRow() {
			private Cell cell = new MockCell() {
				private Date value = testValue;
				@Override
				public CellType getCellType() {
					return CellType.NUMERIC;
				}
				@Override
				public Date getDateCellValue() {
					return value;
				}
				@Override
				public void setCellValue(Date value) {
					this.value = value;
				}
			};
			
			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
		};
	}

	private Row createMockDateRow(String testValue) {
		return new MockRow() {
			private Cell cell = new MockCell() {
				private String value = testValue;
				@Override
				public CellType getCellType() {
					return CellType.STRING;
				}
				@Override
				public String getStringCellValue() {
					return value;
				}
				@Override
				public void setCellValue(String value) {
					this.value = value;
				}
			};

			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
		};
	}

	private Row createMockDateRow(boolean testValue) {
		return new MockRow() {
			private Cell cell = new MockCell() {
				private boolean value = testValue;
				@Override
				public CellType getCellType() {
					return CellType.BOOLEAN;
				}
				@Override
				public boolean getBooleanCellValue() {
					return value;
				}
				@Override
				public void setCellValue(boolean value) {
					this.value = value;
				}
			};
			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
		};
	}

	private Row createMockDateRow() {
		return new MockRow() {
			private Cell cell = null;
			@Override
			public Cell getCell(int cellnum) {
				return cell;
			}
			@Override
			public Cell createCell(int column, CellType type) {
				if (type == CellType.NUMERIC) {
					return cell = new MockCell() {
						private Date value = null;
						@Override
						public CellType getCellType() {
							return CellType.NUMERIC;
						}
						@Override
						public Date getDateCellValue() {
							return value;
						}
						@Override
						public void setCellValue(Date value) {
							this.value = value;
						}
					};
				} else {
					throw new IllegalArgumentException("Invalid CellType '" + type.toString() + "'.");
				}
			}
		};
	}


	private static class MockRow implements Row {

		@Override
		public Iterator<Cell> iterator() {
			throw new UnsupportedOperationException();
		}

		@Override
		public Cell createCell(int column) {
			throw new UnsupportedOperationException();
		}

		@Override
		public Cell createCell(int column, CellType type) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void removeCell(Cell cell) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setRowNum(int rowNum) {
			throw new UnsupportedOperationException();
		}

		@Override
		public int getRowNum() {
			throw new UnsupportedOperationException();
		}

		@Override
		public Cell getCell(int cellnum) {
			throw new UnsupportedOperationException();
		}

		@Override
		public Cell getCell(int cellnum, MissingCellPolicy policy) {
			throw new UnsupportedOperationException();
		}

		@Override
		public short getFirstCellNum() {
			throw new UnsupportedOperationException();
		}

		@Override
		public short getLastCellNum() {
			throw new UnsupportedOperationException();
		}

		@Override
		public int getPhysicalNumberOfCells() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setHeight(short height) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setZeroHeight(boolean zHeight) {
			throw new UnsupportedOperationException();
		}

		@Override
		public boolean getZeroHeight() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setHeightInPoints(float height) {
			throw new UnsupportedOperationException();
		}

		@Override
		public short getHeight() {
			throw new UnsupportedOperationException();
		}

		@Override
		public float getHeightInPoints() {
			throw new UnsupportedOperationException();
		}

		@Override
		public boolean isFormatted() {
			throw new UnsupportedOperationException();
		}

		@Override
		public CellStyle getRowStyle() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setRowStyle(CellStyle style) {
			throw new UnsupportedOperationException();
		}

		@Override
		public Iterator<Cell> cellIterator() {
			throw new UnsupportedOperationException();
		}

		@Override
		public Sheet getSheet() {
			throw new UnsupportedOperationException();
		}

		@Override
		public int getOutlineLevel() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
			throw new UnsupportedOperationException();
		}
	}

	private static class MockCell implements Cell {

		@Override
		public int getColumnIndex() {
			throw new UnsupportedOperationException();
		}

		@Override
		public int getRowIndex() {
			throw new UnsupportedOperationException();
		}

		@Override
		public Sheet getSheet() {
			throw new UnsupportedOperationException();
		}

		@Override
		public Row getRow() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellType(CellType cellType) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setBlank() {
			throw new UnsupportedOperationException();
		}

		@Override
		public CellType getCellType() {
			throw new UnsupportedOperationException();
		}

		@Override
		public CellType getCachedFormulaResultType() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellValue(double value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellValue(Date value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellValue(Calendar value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellValue(RichTextString value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellValue(String value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellFormula(String formula) throws FormulaParseException, IllegalStateException {
			throw new UnsupportedOperationException();
		}

		@Override
		public void removeFormula() throws IllegalStateException {
			throw new UnsupportedOperationException();
		}

		@Override
		public String getCellFormula() {
			throw new UnsupportedOperationException();
		}

		@Override
		public double getNumericCellValue() {
			throw new UnsupportedOperationException();
		}

		@Override
		public Date getDateCellValue() {
			throw new UnsupportedOperationException();
		}

		@Override
		public RichTextString getRichStringCellValue() {
			throw new UnsupportedOperationException();
		}

		@Override
		public String getStringCellValue() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellValue(boolean value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellErrorValue(byte value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public boolean getBooleanCellValue() {
			throw new UnsupportedOperationException();
		}

		@Override
		public byte getErrorCellValue() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellStyle(CellStyle style) {
			throw new UnsupportedOperationException();
		}

		@Override
		public CellStyle getCellStyle() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setAsActiveCell() {
			throw new UnsupportedOperationException();
		}

		@Override
		public CellAddress getAddress() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellComment(Comment comment) {
			throw new UnsupportedOperationException();
		}

		@Override
		public Comment getCellComment() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void removeCellComment() {
			throw new UnsupportedOperationException();
		}

		@Override
		public Hyperlink getHyperlink() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setHyperlink(Hyperlink link) {
			throw new UnsupportedOperationException();
		}

		@Override
		public void removeHyperlink() {
			throw new UnsupportedOperationException();
		}

		@Override
		public CellRangeAddress getArrayFormulaRange() {
			throw new UnsupportedOperationException();
		}

		@Override
		public boolean isPartOfArrayFormulaGroup() {
			throw new UnsupportedOperationException();
		}

		@Override
		public void setCellValue(LocalDateTime value) {
			throw new UnsupportedOperationException();
		}

		@Override
		public LocalDateTime getLocalDateTimeCellValue() {
			throw new UnsupportedOperationException();
		}
		
	}
}
