package net.sf.jett.test;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.model.CellStyleCache;
import net.sf.jett.model.FontCache;
import net.sf.jett.model.ExcelColor;
import net.sf.jett.util.SheetUtil;

/**
 * This JUnit Test class directly tests the <code>CellStyleCache</code> and the
 * <code>FontCache</code>.
 *
 * @author Randy Gettman
 * @since 0.5.0
 */
public class StyleFontCacheTest extends TestCase
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
        String excelNameBase = getExcelNameBase();
        specificTest(TEMPLATES_DIR + excelNameBase + TEMPLATE_SUFFIX + XLS_EXT);
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
        String excelNameBase = getExcelNameBase();
        specificTest(TEMPLATES_DIR + excelNameBase + TEMPLATE_SUFFIX + XLSX_EXT);
    }

    /**
     * Returns the Excel name base for the template and resultant spreadsheets
     * for this test.
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "StyleFontCache";
    }

    /**
     * No actual validation is done here.  For this test, all the work and all
     * the tests are done in the "specificTest" method.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook) { }

    /**
     * Run the test on an Excel spreadsheet defined by name.  This does NOT
     * perform any transformation; this reads in the <code>Workbook</code>, then
     * directly tests the caches.
     * @param inFilename The input filename.
     * @throws IOException If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     */
    private void specificTest(String inFilename) throws IOException, InvalidFormatException
    {
        try (InputStream fileIn = new BufferedInputStream(new FileInputStream(inFilename)))
        {
            Workbook workbook = WorkbookFactory.create(fileIn);
            testCellStyleCache(workbook);
            testFontCache(workbook);
        }
    }

    /**
     * Test the <code>CellStyleCache</code> on the given <code>Workbook</code>.
     * @param workbook The <code>Workbook</code>.
     */
    private void testCellStyleCache(Workbook workbook)
    {
        Sheet sCellStyleCache = workbook.getSheetAt(0);
        CellStyleCache csCache = new CellStyleCache(workbook);
        CellStyle cs = sCellStyleCache.getRow(0).getCell(0).getCellStyle();
        Font fNormal = workbook.getFontAt(cs.getFontIndex());

        int numCellStyles = csCache.getNumEntries();
        // Apparently, Excel HSSF (.xls) has a large number of built-in styles
        // that somehow have the same string representation.
        // Don't ensure that the number of cell styles is equal to the number of
        // cache entries.
        //assertEquals(numCellStyles, csCache.getNumEntries());

        // Defaults.
        HorizontalAlignment alignment = cs.getAlignment();
        BorderStyle borderBottom = cs.getBorderBottom();
        BorderStyle borderLeft = cs.getBorderLeft();
        BorderStyle borderRight = cs.getBorderRight();
        BorderStyle borderTop = cs.getBorderTop();
        String dataFormat = cs.getDataFormatString();
        Color fillBackgroundColor = cs.getFillBackgroundColorColor();
        Color fillForegroundColor = cs.getFillForegroundColorColor();
        FillPatternType fillPattern = cs.getFillPattern();
        boolean hidden = cs.getHidden();
        short indention = cs.getIndention();
        boolean locked = cs.getLocked();
        short rotationDegrees = cs.getRotation();
        VerticalAlignment verticalAlignment = cs.getVerticalAlignment();
        boolean wrapText = cs.getWrapText();
        // Don't bother actually getting it from the CellStyle here, which would
        // involve HSSF/XSSF-specific processing.
        Color bottomBorderColor = null;
        Color leftBorderColor = null;
        Color rightBorderColor = null;
        Color topBorderColor = null;
        // Font properties (shouldn't change at all for CellStyle tests).
        boolean fontBoldweight = fNormal.getBold();
        int fontCharset = fNormal.getCharSet();
        Color fontColor;
        if (workbook instanceof HSSFWorkbook)
        {
            fontColor = ExcelColor.getHssfColorByIndex(fNormal.getColor());
        }
        else
        {
            // XSSFWorkbook
            // See StyleTag.java comments for why we're creating a new XSSFColor
            // instead of just using the font-supplied XSSFColor.
            fontColor = new XSSFColor(((XSSFFont) fNormal).getXSSFColor().getRGB(),new DefaultIndexedColorMap());
        }
        short fontHeightInPoints = fNormal.getFontHeightInPoints();
        String fontName = fNormal.getFontName();
        boolean fontItalic = fNormal.getItalic();
        boolean fontStrikeout = fNormal.getStrikeout();
        short fontTypeOffset = fNormal.getTypeOffset();
        byte fontUnderline = fNormal.getUnderline();

        // Alignment.
        // Expect a cache hit.
        CellStyle cached = csCache.retrieveCellStyle(fontBoldweight, fontItalic, fontColor, fontName, fontHeightInPoints,
                alignment, borderBottom, borderLeft, borderRight, borderTop, dataFormat, fontUnderline, fontStrikeout,
                wrapText, fillBackgroundColor, fillForegroundColor, fillPattern, verticalAlignment, indention, rotationDegrees,
                bottomBorderColor, leftBorderColor, rightBorderColor, topBorderColor, fontCharset, fontTypeOffset,
                locked, hidden);
        assertNotNull(cached);
        HorizontalAlignment newAlignment = HorizontalAlignment.RIGHT;
        // Expect a cache miss.
        CellStyle notCached = csCache.retrieveCellStyle(fontBoldweight, fontItalic, fontColor, fontName, fontHeightInPoints,
         /* changed */ newAlignment, borderBottom, borderLeft, borderRight, borderTop, dataFormat, fontUnderline, fontStrikeout,
                wrapText, fillBackgroundColor, fillForegroundColor, fillPattern, verticalAlignment, indention, rotationDegrees,
                bottomBorderColor, leftBorderColor, rightBorderColor, topBorderColor, fontCharset, fontTypeOffset,
                locked, hidden);
        assertNull(notCached);
        CellStyle newStyle = SheetUtil.createCellStyle(workbook, newAlignment, borderBottom, borderLeft, borderRight,
                borderTop, dataFormat, wrapText, fillBackgroundColor, fillForegroundColor, fillPattern, verticalAlignment,
                indention, rotationDegrees, bottomBorderColor, leftBorderColor, rightBorderColor, topBorderColor,
                locked, hidden);
        csCache.cacheCellStyle(newStyle);
        numCellStyles++;
        assertEquals(numCellStyles, csCache.getNumEntries());
        CellStyle nowCached = csCache.retrieveCellStyle(fontBoldweight, fontItalic, fontColor, fontName, fontHeightInPoints,
         /* changed */ newAlignment, borderBottom, borderLeft, borderRight, borderTop, dataFormat, fontUnderline, fontStrikeout,
                wrapText, fillBackgroundColor, fillForegroundColor, fillPattern, verticalAlignment, indention, rotationDegrees,
                bottomBorderColor, leftBorderColor, rightBorderColor, topBorderColor, fontCharset, fontTypeOffset,
                locked, hidden);
        assertNotNull(nowCached);
    }

    /**
     * Test the <code>FontCache</code> on the given <code>Workbook</code>.
     * @param workbook The <code>Workbook</code>.
     */
    private void testFontCache(Workbook workbook)
    {
        Sheet sFontCache = workbook.getSheetAt(1);
        FontCache fCache = new FontCache(workbook);
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
}
