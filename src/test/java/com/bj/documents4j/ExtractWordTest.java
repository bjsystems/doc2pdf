package com.bj.documents4j;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class ExtractWordTest {

	@Test
	public void extract_워드파일_변환_Test() {
		ExtractWord ew = new ExtractWord();
		boolean result = ew.extract();
		
		assertEquals(true, result);
	}
}
