package com._4point.utils.simpleexcel.api;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com._4point.utils.simpleexcel.api.Xlsx.XlsxException;

class XlsxExceptionTest {

	private static final String TEST_MESSAGE = "Test message text.";
	private static final Throwable TEST_EXCEPTION = new IllegalStateException();

	@BeforeAll
	static void setUpBeforeClass() throws Exception {
	}

	@BeforeEach
	void setUp() throws Exception {
	}

	@Test
	void testXlsxExceptionStringThrowable() {
		XlsxException underTest = new XlsxException(TEST_MESSAGE, TEST_EXCEPTION);
		checkMessage(underTest);
		checkCause(underTest);
	}

	@Test
	void testXlsxExceptionString() {
		XlsxException underTest = new XlsxException(TEST_MESSAGE);
		checkMessage(underTest);
	}

	@Test
	void testXlsxExceptionThrowable() {
		XlsxException underTest = new XlsxException(TEST_EXCEPTION);
		checkCause(underTest);
	}

	private void checkMessage(XlsxException underTest) {
		final String message = underTest.getMessage();
		assertNotNull(message);
		assertEquals(TEST_MESSAGE, message);
	}

	private void checkCause(XlsxException underTest) {
		final Throwable cause = underTest.getCause();
		assertNotNull(cause);
		assertEquals(TEST_EXCEPTION, cause);
	}
}
