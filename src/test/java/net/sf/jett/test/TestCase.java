package net.sf.jett.test;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import static org.junit.Assert.*;

import net.sf.jett.transform.ExcelTransformer;

/**
 * A <code>TestCase</code> is the superclass for all JETT JUnit test classes.
 *
 * @author Randy Gettman
 */
public abstract class TestCase
{

    /**
     * Standard extremely small delta value to satisfy
     * <code>Assert.assertEquals(double, double, double)</code>.
     */
    public static final double DELTA = 0.00000000001;

    protected static final String TEMPLATES_DIR = "templates/";
    protected static final String OUTPUT_DIR = "output/";
    protected static final String TEMPLATE_SUFFIX = "Template";
    protected static final String OUTPUT_SUFFIX = "Result";
    protected static final String XLS_EXT = ".xls";
    protected static final String XLSX_EXT = ".xlsx";
    protected static final String FILES_IND = "Files";

    protected boolean amISetup = false;
    protected Map<String, Object> myBeansMap;
    protected List<Map<String, Object>> myListOfBeansMaps;
    protected List<String> myTemplateSheetNames;
    protected List<String> myResultSheetNames;

    /**
     * Tests the .xls template spreadsheet.  This is meant to have the
     * <code>@Test</code> annotation in all concrete subclasses.  Also, each
     * concrete subclass should simply call <code>super.testXls</code>.
     *
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     */
    public void testXls() throws IOException, InvalidFormatException
    {
        File fOutputDir = new File(OUTPUT_DIR);
        if (!fOutputDir.exists() && !fOutputDir.mkdirs())
        {
            throw new RuntimeException("Couldn't create output directory: " + OUTPUT_DIR);
        }
        String excelNameBase = getExcelNameBase();
        genericTest(TEMPLATES_DIR + excelNameBase + TEMPLATE_SUFFIX + XLS_EXT,
                OUTPUT_DIR + excelNameBase + OUTPUT_SUFFIX + XLS_EXT);
    }

    /**
     * Tests the .xlsx template spreadsheet.  This is meant to have the
     * <code>@Test</code> annotation in all concrete subclasses.  Also, each
     * concrete subclass should simply call <code>super.testXlsx</code>.
     *
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     */
    public void testXlsx() throws IOException, InvalidFormatException
    {
        File fOutputDir = new File(OUTPUT_DIR);
        if (!fOutputDir.exists() && !fOutputDir.mkdirs())
        {
            throw new RuntimeException("Couldn't create output directory: " + OUTPUT_DIR);
        }
        String excelNameBase = getExcelNameBase();
        genericTest(TEMPLATES_DIR + excelNameBase + TEMPLATE_SUFFIX + XLSX_EXT,
                OUTPUT_DIR + excelNameBase + OUTPUT_SUFFIX + XLSX_EXT);
    }

    /**
     * Tests the .xls template spreadsheet.  This is meant to have the
     * <code>@Test</code> annotation in all concrete subclasses.  Also, each
     * concrete subclass should simply call <code>super.testXlsFiles</code>.
     *
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     * @since 0.2.0
     */
    public void testXlsFiles() throws IOException, InvalidFormatException
    {
        File fOutputDir = new File(OUTPUT_DIR);
        if (!fOutputDir.exists() && !fOutputDir.mkdirs())
        {
            throw new RuntimeException("Couldn't create output directory: " + OUTPUT_DIR);
        }
        String excelNameBase = getExcelNameBase();
        genericTestFiles(TEMPLATES_DIR + excelNameBase + TEMPLATE_SUFFIX + XLS_EXT,
                OUTPUT_DIR + excelNameBase + FILES_IND + OUTPUT_SUFFIX + XLS_EXT);
    }

    /**
     * Tests the .xlsx template spreadsheet.  This is meant to have the
     * <code>@Test</code> annotation in all concrete subclasses.  Also, each
     * concrete subclass should simply call <code>super.testXlsFiles</code>.
     *
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     * @since 0.2.0
     */
    public void testXlsxFiles() throws IOException, InvalidFormatException
    {
        File fOutputDir = new File(OUTPUT_DIR);
        if (!fOutputDir.exists() && !fOutputDir.mkdirs())
        {
            throw new RuntimeException("Couldn't create output directory: " + OUTPUT_DIR);
        }
        String excelNameBase = getExcelNameBase();
        genericTestFiles(TEMPLATES_DIR + excelNameBase + TEMPLATE_SUFFIX + XLSX_EXT,
                OUTPUT_DIR + excelNameBase + FILES_IND + OUTPUT_SUFFIX + XLSX_EXT);
    }

    /**
     * Runs the actual test on an Excel template spreadsheet.
     *
     * @param inFilename  The input filename.
     * @param outFilename The output filename.
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     */
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
                workbook = transformer.transform(
                        fileIn, myTemplateSheetNames, myResultSheetNames, myListOfBeansMaps);
            }
            else
            {
                if (!amISetup)
                {
                    myBeansMap = getBeansMap();
                    amISetup = true;
                }
                assertNotNull(myBeansMap);
                workbook = transformer.transform(fileIn, myBeansMap);
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
     * Runs the actual test on Excel files, from input template filename to
     * output filename.
     *
     * @param inFilename  The input filename.
     * @param outFilename The output filename.
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     * @since 0.2.0
     */
    private void genericTestFiles(String inFilename, String outFilename)
            throws IOException, InvalidFormatException
    {
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
            transformer.transform(inFilename, outFilename,
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
            transformer.transform(inFilename, outFilename, myBeansMap);
        }

        try (InputStream fileIn = new BufferedInputStream(new FileInputStream(outFilename)))
        {
            Workbook workbook = WorkbookFactory.create(fileIn);
            check(workbook);
        }
    }

    /**
     * Returns the Excel name base for the template and resultant spreadsheets
     * for this test.
     *
     * @return The Excel name base for this test.
     */
    protected abstract String getExcelNameBase();

    /**
     * Call certain setup-related methods on the <code>ExcelTransformer</code>
     * before template sheet transformation.
     *
     * @param transformer The <code>ExcelTransformer</code> that will transform
     *                    the template worksheet(s).
     */
    protected void setupTransformer(ExcelTransformer transformer)
    {
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.  Helper methods are available in the <code>TestUtility</code>
     * class.
     *
     * @param workbook A <code>Workbook</code>.
     * @see TestUtility
     */
    protected abstract void check(Workbook workbook);

    /**
     * Determines whether this test uses a single map of beans, or if it uses
     * multiple maps of beans along with template sheet names and resultant
     * sheet names.
     *
     * @return <code>true</code> if this test uses multiple bean maps, or
     * <code>false</code> if this test uses a single map of beans.
     */
    protected abstract boolean isMultipleBeans();

    /**
     * For multiple beans map tests, return the <code>List</code> of template
     * sheet names.
     *
     * @return A <code>List</code> of template sheet names.
     */
    protected List<String> getListOfTemplateSheetNames()
    {
        return null;
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of result
     * sheet names.
     *
     * @return A <code>List</code> of result sheet names.
     */
    protected List<String> getListOfResultSheetNames()
    {
        return null;
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of beans maps,
     * which map bean names to bean values for each corresponding sheet.
     *
     * @return A <code>List</code> of <code>Maps</code> of bean names to bean
     * values.
     */
    protected List<Map<String, Object>> getListOfBeansMaps()
    {
        return null;
    }

    /**
     * For single beans map tests, return the <code>Map</code> of bean names to
     * bean values.
     *
     * @return A <code>Map</code> of bean names to bean values.
     */
    protected Map<String, Object> getBeansMap()
    {
        return null;
    }
}