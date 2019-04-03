package net.sf.jett.test;

import java.io.IOException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import static org.junit.Assert.*;

import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the ability of <code>ExcelTransformer</code> to
 * supply different bean maps to <strong>non-consecutive</strong> cloned
 * <code>Sheets</code>.
 *
 * @author Randy Gettman
 * @since 0.7.0
 */
public class MultipleNonConsecutiveTest extends TestCase
{
    /**
     * Overriding only to get the <code>Workbook</code> object and manipulate it
     * <em>before</em> transformation.
     *
     * @param inFilename  The input filename.
     * @param outFilename The output filename.
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     * @since 0.7.0
     */
    @Override
    protected void genericTest(String inFilename, String outFilename)
            throws IOException, InvalidFormatException
    {
        try (FileOutputStream fileOut = new FileOutputStream(outFilename);
             InputStream fileIn = new BufferedInputStream(new FileInputStream(inFilename)))
        {
            Workbook workbook;
            ExcelTransformer transformer = new ExcelTransformer();
            setupTransformer(transformer);
            if (isMultipleBeans())
            {
                if (!amISetup)
                {
                    myTemplateSheetNames = getListOfTemplateSheetNames();
                    myResultSheetNames = getListOfResultSheetNames();
                    myListOfBeansMaps = getListOfBeansMaps();
                    amISetup = true;
                }
                assertNotNull(myTemplateSheetNames);
                assertNotNull(myResultSheetNames);
                assertNotNull(myListOfBeansMaps);
                workbook = WorkbookFactory.create(fileIn);
                beforeTransformation(workbook);
                transformer.transform(workbook,
                        myTemplateSheetNames, myResultSheetNames, myListOfBeansMaps);
            }
            else
            {
                if (!amISetup)
                {
                    myBeansMap = getBeansMap();
                    amISetup = true;
                }
                assertNotNull(myBeansMap);
                workbook = WorkbookFactory.create(fileIn);
                beforeTransformation(workbook);
                transformer.transform(workbook, myBeansMap);
            }

            // Becomes invalid after write().
            Error error = null;
            RuntimeException exception = null;
            try
            {
                if (!(workbook instanceof HSSFWorkbook))
                    check(workbook);
            }
            catch (RuntimeException e)
            {
                exception = e;
            }
            catch (Error e)
            {
                error = e;
            }

            workbook.write(fileOut);
            fileOut.close();

            if (error != null)
            {
                error.printStackTrace();
                fail();
            }
            if (exception != null)
            {
                exception.printStackTrace();
                throw exception;
            }

            // Check HSSF after writing to see errors.
            if (workbook instanceof HSSFWorkbook)
                check(workbook);
        }
    }

    /**
     * Tests the .xls template spreadsheet.
     *
     * @throws java.io.IOException                                        If an I/O error occurs.
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
     *
     * @throws java.io.IOException                                        If an I/O error occurs.
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
     *
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "MultipleNonConsecutive";
    }

    /**
     * A chance to manipulate the <code>Workbook</code> <em>before</em> it is
     * transformed.
     *
     * @param workbook The template <code>Workbook</code>.
     */
    private void beforeTransformation(Workbook workbook)
    {
        Sheet symbols = workbook.getSheetAt(1);
        PrintSetup ps = symbols.getPrintSetup();

        // These settings were all found not to be copied upon a call to
        // "cloneSheet".  Test them by setting them in the template sheet.
        // The "checkSheet" method will check all resultant sheets to see if they
        // retain these settings.
        symbols.setRepeatingColumns(CellRangeAddress.valueOf("A:A"));
        symbols.setRepeatingRows(CellRangeAddress.valueOf("1:1"));

        ps.setCopies((short) 2);
        ps.setDraft(true);
        ps.setFitHeight((short) 2);
        ps.setFitWidth((short) 2);
        ps.setHResolution((short) 300);
        ps.setLandscape(true);
        ps.setLeftToRight(true);
        ps.setNoColor(true);
        ps.setNotes(true);
        ps.setPageStart((short) 2);
        ps.setPaperSize(PrintSetup.LEGAL_PAPERSIZE);
        ps.setScale((short) 101);
        ps.setUsePage(true);
        ps.setValidSettings(false);
        ps.setVResolution((short) 300);
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     *
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet g1 = workbook.getSheetAt(0);
        assertEquals("G1", g1.getSheetName());
        assertEquals("!", TestUtility.getStringCellValue(g1, 1, 0));
        checkSheet(g1);

        Sheet b1 = workbook.getSheetAt(1);
        assertEquals("B1", b1.getSheetName());
        assertEquals("0", TestUtility.getStringCellValue(b1, 1, 0));

        Sheet r1 = workbook.getSheetAt(2);
        assertEquals("R1", r1.getSheetName());
        assertEquals("A", TestUtility.getStringCellValue(r1, 1, 0));

        Sheet g2 = workbook.getSheetAt(3);
        assertEquals("G2", g2.getSheetName());
        assertEquals("!", TestUtility.getStringCellValue(g2, 1, 0));
        checkSheet(g2);

        Sheet g3 = workbook.getSheetAt(4);
        assertEquals("G3", g3.getSheetName());
        assertEquals("!", TestUtility.getStringCellValue(g3, 1, 0));
        checkSheet(g3);

        Sheet b2 = workbook.getSheetAt(5);
        assertEquals("B2", b2.getSheetName());
        assertEquals("0", TestUtility.getStringCellValue(b2, 1, 0));

        Sheet b3 = workbook.getSheetAt(6);
        assertEquals("B3", b3.getSheetName());
        assertEquals("0", TestUtility.getStringCellValue(b3, 1, 0));

        Sheet g4 = workbook.getSheetAt(7);
        assertEquals("G4", g4.getSheetName());
        assertEquals("!", TestUtility.getStringCellValue(g4, 1, 0));
        checkSheet(g4);

        Sheet b4 = workbook.getSheetAt(8);
        assertEquals("B4", b4.getSheetName());
        assertEquals("0", TestUtility.getStringCellValue(b4, 1, 0));

        Sheet r2 = workbook.getSheetAt(9);
        assertEquals("R2", r2.getSheetName());
        assertEquals("A", TestUtility.getStringCellValue(r2, 1, 0));

        Sheet g5 = workbook.getSheetAt(10);
        assertEquals("G5", g5.getSheetName());
        assertEquals("!", TestUtility.getStringCellValue(g5, 1, 0));
        checkSheet(g5);

        Sheet b5 = workbook.getSheetAt(11);
        assertEquals("B5", b5.getSheetName());
        assertEquals("0", TestUtility.getStringCellValue(b5, 1, 0));
    }

    /**
     * Checks the given <code>Sheet</code> if JETT has copied the settings to
     * the cloned <code>Sheet</code> that are not automatically copied upon a
     * call to <code>cloneSheet</code>.
     *
     * @param sheet A resultant <code>Sheet</code>.
     */
    private void checkSheet(Sheet sheet)
    {
        PrintSetup ps = sheet.getPrintSetup();

        // These settings were all found not to be copied upon a call to
        // "cloneSheet".  Test them by setting them in the template sheet.
        // The "checkSheet" method will check all resultant sheets to see if they
        // retain these settings.
        //assertEquals("org.apache.poi.ss.util.CellRangeAddress [A:A]", sheet.getRepeatingColumns().toString());
        //assertEquals("org.apache.poi.ss.util.CellRangeAddress [1:1]", sheet.getRepeatingRows().toString());

        assertEquals(2, ps.getCopies());
        assertTrue(ps.getDraft());
        assertEquals(2, ps.getFitHeight());
        assertEquals(2, ps.getFitWidth());
        assertEquals(300, ps.getHResolution());
        assertTrue(ps.getLandscape());
        assertTrue(ps.getLeftToRight());
        assertTrue(ps.getNoColor());
        assertTrue(ps.getNotes());
        assertEquals(2, ps.getPageStart());
        assertEquals(PrintSetup.LEGAL_PAPERSIZE, ps.getPaperSize());
        assertEquals(101, ps.getScale());
        assertTrue(ps.getUsePage());
        assertFalse(ps.getValidSettings());
        assertEquals(300, ps.getVResolution());
    }

    /**
     * This test is a multiple beans map test.
     *
     * @return <code>true</code>.
     */
    @Override
    protected boolean isMultipleBeans()
    {
        return true;
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of template
     * sheet names.
     *
     * @return A <code>List</code> of template sheet names.
     */
    @Override
    protected List<String> getListOfTemplateSheetNames()
    {
        return Arrays.asList("Green", "Blue", "Red", "Green", "Green", "Blue", "Blue", "Green", "Blue", "Red", "Green", "Blue");
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of result
     * sheet names.
     *
     * @return A <code>List</code> of result sheet names.
     */
    @Override
    protected List<String> getListOfResultSheetNames()
    {
        return Arrays.asList("G1", "B1", "R1", "G2", "G3", "B2", "B3", "G4", "B4", "R2", "G5", "B5");
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of beans maps,
     * which map bean names to bean values for each corresponding sheet.
     *
     * @return A <code>List</code> of <code>Maps</code> of bean names to bean
     * values.
     */
    @Override
    protected List<Map<String, Object>> getListOfBeansMaps()
    {
        Map<String, Object> letters = new HashMap<>();
        letters.put("letters", Arrays.asList("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
                "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"));
        Map<String, Object> symbols = new HashMap<>();
        symbols.put("symbols", Arrays.asList("!", "@", "#", "$", "%", "^", "&", "*", "(", ")"));
        Map<String, Object> numbers = new HashMap<>();
        numbers.put("numbers", Arrays.asList("0", "1", "2", "3", "4", "5", "6", "7", "8", "9"));
        List<Map<String, Object>> beansMaps = new ArrayList<>();
        beansMaps.add(symbols);
        beansMaps.add(numbers);
        beansMaps.add(letters);
        beansMaps.add(symbols);
        beansMaps.add(symbols);
        beansMaps.add(numbers);
        beansMaps.add(numbers);
        beansMaps.add(symbols);
        beansMaps.add(numbers);
        beansMaps.add(letters);
        beansMaps.add(symbols);
        beansMaps.add(numbers);
        return beansMaps;
    }
}