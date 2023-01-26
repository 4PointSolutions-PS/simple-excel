package com._4point.utils.simpleexcel.api;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;

public interface Xlsx {

	public Optional<Worksheet> getSheet(String worksheetName);
	public Worksheet createSheet(String worksheetName);
	public void save(Path filePath) throws IOException;

	@SuppressWarnings("serial")
	public static class XlsxException extends Exception {
		
		public XlsxException(String message, Throwable cause) {
			super(message, cause);
		}
	
		public XlsxException(String message) {
			super(message);
		}
	
		public XlsxException(Throwable cause) {
			super(cause);
		}		
	}
}
