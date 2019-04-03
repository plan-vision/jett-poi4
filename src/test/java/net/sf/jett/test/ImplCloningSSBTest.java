package net.sf.jett.test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

/**
 * This JUnit Test class tests the implicit cloning feature of JETT, with
  * "sheet specific beans" transformation.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class ImplCloningSSBTest extends ImplicitCloningTest
{
    /**
     * Tests the .xls template spreadsheet.
     *
     * @throws IOException If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
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
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
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
        return "ImplCloningSSB";
    }

    /**
     * This test is a single map test.
     *
     * @return <code>false</code>.
     */
    @Override
    protected boolean isMultipleBeans()
    {
        return true;
    }


    /**
     * For multiple beans map tests, return the <code>List</code> of template
     * sheet names.
     * @return A <code>List</code> of template sheet names.
     */
    @Override
    protected List<String> getListOfTemplateSheetNames()
    {
       return Arrays.asList("Static1", "${dvs.name}$@i=n;l=10;v=s;r=DNE", "Static2", "${dvs.name}$@l=0", "Static3", "${dvs.name}$@l=1");
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of result
     * sheet names.
     * @return A <code>List</code> of result sheet names.
     */
    @Override
    protected List<String> getListOfResultSheetNames()
    {
        return Arrays.asList("Static1", "${dvs.name}$@i=n;l=10;v=s;r=DNE", "Static2", "${dvs.name}$@l=0", "Static3", "${dvs.name}$@l=1");
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of beans maps,
     * which map bean names to bean values for each corresponding sheet.
     * @return A <code>List</code> of <code>Maps</code> of bean names to bean
     *    values.
     */
    @Override
    protected List<Map<String, Object>> getListOfBeansMaps()
    {
        Map<String, Object> dvs = TestUtility.getDivisionData();
        // Space is at a premium in sheet names -- 31 characters.
        dvs.put("dvs", dvs.get("divisionsList"));
        dvs.remove("divisionsList");
        List<Map<String, Object>> beansList = new ArrayList<>();
        for (int i = 0; i < 6; i++)
        {
            if (i % 2 > 0)
            {
                beansList.add(dvs);
            }
            else
            {
                beansList.add(new HashMap<String, Object>());
            }
        }
        return beansList;
    }
}
