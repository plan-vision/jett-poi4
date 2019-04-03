package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Font;
import org.junit.Test;
import static org.junit.Assert.*;
import net.sf.jett.util.RichTextStringUtil;

/**
 * This JUnit Test class tests the evaluation of the "comment" tag.
 *
 * @author Randy Gettman
 */
public class CommentTagTest extends TestCase
{
    /**
     * Tests the .xls template spreadsheet.
     * @throws java.io.IOException If an I/O error occurs.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If the input spreadsheet is invalid.
     */
    @Override
    @Test
    public void testXls() throws IOException, InvalidFormatException
    {
        super.testXls();
    }

    /**
     * Tests the .xlsx template spreadsheet.
     * @throws java.io.IOException If an I/O error occurs.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If the input spreadsheet is invalid.
     */
    @Override
    @Test
    public void testXlsx() throws IOException, InvalidFormatException
    {
        super.testXlsx();
    }

    /**
     * Returns the Excel name base for the template and resultant spreadsheets
     * for this test.
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "CommentTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet sComment = workbook.getSheetAt(0);
        Comment comment;

        assertEquals("City: Boston", TestUtility.getStringCellValue(sComment, 2, 0));
        RichTextString rts = TestUtility.getRichTextStringCellValue(sComment, 2, 0);
        Font font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
        assertTrue(font.getBold());
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 4), workbook);
        assertTrue(font.getBold());
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 5), workbook);
        assertTrue(font == null || !font.getBold());
        comment = TestUtility.getComment(sComment, 2, 0);
        assertNotNull(comment);
        assertFalse(comment.isVisible());
        assertEquals("Team Name: Celtics", comment.getString().getString());
        rts = comment.getString();
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
        assertTrue(font.getBold());
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 9), workbook);
        assertTrue(font.getBold());
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 10), workbook);
        assertTrue(font == null || !font.getBold());

        assertEquals("Atlantic Division", comment.getAuthor());
        assertEquals("City: Toronto", TestUtility.getStringCellValue(sComment, 6, 0));
        comment = TestUtility.getComment(sComment, 6, 0);
        assertNotNull(comment);
        assertFalse(comment.isVisible());
        assertEquals("Team Name: Raptors", comment.getString().getString());
        assertEquals("Atlantic Division", comment.getAuthor());

        assertEquals("City: Los Angeles", TestUtility.getStringCellValue(sComment, 30, 0));
        comment = TestUtility.getComment(sComment, 30, 0);
        assertNotNull(comment);
        assertFalse(comment.isVisible());
        assertEquals("Team Name: Lakers", comment.getString().getString());
        assertEquals("Pacific Division", comment.getAuthor());
        assertEquals("City: Sacramento", TestUtility.getStringCellValue(sComment, 34, 0));
        comment = TestUtility.getComment(sComment, 34, 0);
        assertNotNull(comment);
        assertFalse(comment.isVisible());
        assertEquals("Team Name: Kings", comment.getString().getString());
        assertEquals("Pacific Division", comment.getAuthor());

        assertEquals("City: Harlem", TestUtility.getStringCellValue(sComment, 46, 0));
        comment = TestUtility.getComment(sComment, 46, 0);
        assertNotNull(comment);
        assertTrue(comment.isVisible());
        assertEquals("Team Name: Globetrotters", comment.getString().getString());
        assertEquals("Of Their Own Division", comment.getAuthor());
    }

    /**
     * This test is a single map test.
     * @return <code>false</code>.
     */
    @Override
    protected boolean isMultipleBeans()
    {
        return false;
    }

    /**
     * For single beans map tests, return the <code>Map</code> of bean names to
     * bean values.
     * @return A <code>Map</code> of bean names to bean values.
     */
    @Override
    protected Map<String, Object> getBeansMap()
    {
        return TestUtility.getDivisionData();
    }
}