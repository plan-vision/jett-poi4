package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.junit.Test;
import static org.junit.Assert.*;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import net.sf.jett.util.RichTextStringUtil;

/**
 * This JUnit Test class tests the evaluation of the "hyperlink" tag.
 *
 * @author Randy Gettman
 */
public class HyperlinkTagTest extends TestCase {
	/**
	 * Tests the .xls template spreadsheet.
	 * 
	 * @throws java.io.IOException If an I/O error occurs.
	 * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If the
	 *         input spreadsheet is invalid.
	 */
	@Override
	@Test
	public void testXls() throws IOException, InvalidFormatException {
		super.testXls();
	}

	/**
	 * Tests the .xlsx template spreadsheet.
	 * 
	 * @throws IOException            If an I/O error occurs.
	 * @throws InvalidFormatException If the input spreadsheet is invalid.
	 */
	@Override
	@Test
	public void testXlsx() throws IOException, InvalidFormatException {
		super.testXlsx();
	}

	/**
	 * Returns the Excel name base for the template and resultant spreadsheets for
	 * this test.
	 * 
	 * @return The Excel name base for this test.
	 */
	@Override
	protected String getExcelNameBase() {
		return "HyperlinkTag";
	}

	/**
	 * Validate the newly created resultant <code>Workbook</code> with JUnit
	 * assertions.
	 * 
	 * @param workbook A <code>Workbook</code>.
	 */
	@Override
	protected void check(Workbook workbook) {
		Sheet sheet = workbook.getSheetAt(0);
		RichTextString value1 = TestUtility.getRichTextStringCellValue(sheet, 1, 0);
		assertNotNull(value1);
		assertEquals("JETT on SourceForge", value1.getString());
		RichTextString rts = TestUtility.getRichTextStringCellValue(sheet, 1, 0);
		Font font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
		assertEquals("0000ff", TestUtility.getFontColorString(workbook, font));
		Hyperlink h = TestUtility.getHyperlink(sheet, 1, 0);
		assertNotNull(h);
		assertEquals(HyperlinkType.URL, h.getType());
		assertEquals("http://jett.sourceforge.net", h.getAddress());

		assertEquals(Font.U_SINGLE, font.getUnderline());
		assertEquals("Email jett-users", TestUtility.getStringCellValue(sheet, 2, 0));
		rts = TestUtility.getRichTextStringCellValue(sheet, 2, 0);
		font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
		assertEquals("0000ff", TestUtility.getFontColorString(workbook, font));
		assertEquals(Font.U_SINGLE, font.getUnderline());
		h = TestUtility.getHyperlink(sheet, 2, 0);
		assertNotNull(h);
		assertEquals(HyperlinkType.EMAIL, h.getType());
		assertEquals("mailto:jett-users@lists.sourceforge.net", h.getAddress());

		assertEquals("Template For This Test (.xlsx)", TestUtility.getStringCellValue(sheet, 3, 0));
		rts = TestUtility.getRichTextStringCellValue(sheet, 3, 0);
		font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
		assertEquals("0000ff", TestUtility.getFontColorString(workbook, font));
		assertEquals(Font.U_SINGLE, font.getUnderline());
		h = TestUtility.getHyperlink(sheet, 3, 0);
		assertNotNull(h);
		assertEquals(HyperlinkType.FILE, h.getType());
		assertEquals("../templates/HyperlinkTagTemplate.xlsx", h.getAddress());

		assertEquals("Intra-spreadsheet Link", TestUtility.getStringCellValue(sheet, 4, 0));
		rts = TestUtility.getRichTextStringCellValue(sheet, 4, 0);
		font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
		assertEquals("0000ff", TestUtility.getFontColorString(workbook, font));
		assertEquals(Font.U_SINGLE, font.getUnderline());
		h = TestUtility.getHyperlink(sheet, 4, 0);
		assertNotNull(h);
		assertEquals(HyperlinkType.DOCUMENT, h.getType());
		assertEquals("'Target Sheet'!B3", h.getAddress());

		assertEquals("Additional Help", TestUtility.getStringCellValue(sheet, 5, 0));
		rts = TestUtility.getRichTextStringCellValue(sheet, 5, 0);
		font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
		assertEquals("0000ff", TestUtility.getFontColorString(workbook, font));
		assertEquals(Font.U_SINGLE, font.getUnderline());
		h = TestUtility.getHyperlink(sheet, 5, 0);
		assertNotNull(h);
		assertEquals(HyperlinkType.URL, h.getType());
		assertEquals("http://www.youtube.com/watch?v=dQw4w9WgXcQ", h.getAddress());

		// Shift
		Sheet shift = workbook.getSheetAt(2);
		for (int r = 1; r <= 10; r++) {
			CellStyle cs = TestUtility.getCellStyle(shift, r, 1);
			font = workbook.getFontAt(cs.getFontIndex());
			assertEquals("0000ff", TestUtility.getFontColorString(workbook, font));
			assertEquals(Font.U_SINGLE, font.getUnderline());
			h = TestUtility.getHyperlink(shift, r, 1);
			assertNotNull(h);
			assertEquals(HyperlinkType.URL, h.getType());
			assertEquals("http://www.example.com/", h.getAddress());
		}

		CellStyle cs = TestUtility.getCellStyle(shift, 11, 1);
		font = workbook.getFontAt(cs.getFontIndex());
		assertEquals("0000ff", TestUtility.getFontColorString(workbook, font));
		assertEquals(Font.U_SINGLE, font.getUnderline());
		h = TestUtility.getHyperlink(shift, 11, 1);
		assertNotNull(h);
		assertEquals(HyperlinkType.URL, h.getType());
		assertEquals("http://jett.sourceforge.net/", h.getAddress());
	}

	/**
	 * This test is a single map test.
	 * 
	 * @return <code>false</code>.
	 */
	@Override
	protected boolean isMultipleBeans() {
		return false;
	}

	/**
	 * For single beans map tests, return the <code>Map</code> of bean names to bean
	 * values.
	 * 
	 * @return A <code>Map</code> of bean names to bean values.
	 */
	@Override
	protected Map<String, Object> getBeansMap() {
		return TestUtility.getHyperlinkData();
	}
}
