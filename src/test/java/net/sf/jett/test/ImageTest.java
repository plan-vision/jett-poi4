package net.sf.jett.test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "comment" tag.
 *
 * @author Randy Gettman
 * @since 0.10.0
 */
public class ImageTest extends TestCase
{
    /**
     * Tests the .xls template spreadsheet.
     *
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
     *
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
     *
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "ImageTag";
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
        List<? extends PictureData> pictures = workbook.getAllPictures();
        List<String> origFilenames = Arrays.asList("templates/1994.hyundai.excel.5263-396x249.png",
                                                   "templates/1994.hyundai.excel.5263-396x249.jpg");
        for (int i = 0; i < origFilenames.size(); i++)
        {
            PictureData pict = pictures.get(i);
            byte[] data = pict.getData();
            try
            {
                InputStream is = new FileInputStream(origFilenames.get(i));
                byte[] orig = IOUtils.toByteArray(is);

                assertArrayEquals("Mismatch on case: " + i, orig, data);
            }
            catch (IOException e)
            {
                fail("Exception on case: " + i + ", " + e.getMessage());
            }
        }
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
        Map<String, Object> beans = new HashMap<>();
        beans.put("width", 4);
        beans.put("height", 6);
        return beans;
    }
}
