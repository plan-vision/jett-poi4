package net.sf.jett.test;

import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import static org.junit.Assert.assertEquals;

/**
 * This JUnit Test class tests the implicit cloning feature of JETT.  It is
 * subclassed by other classes that perform normal and sheet-specific beans
 * testing.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public abstract class ImplicitCloningTest extends TestCase
{
    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     *
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet static1 = workbook.getSheetAt(0);
        // "OfTheirOwn" instead of "Of Their Own"; getFormulaCellValue strips spaces.
        assertEquals("COUNTA(Atlantic!B3:B7,Central!B3:B7,Northwest!B3:B7,'OfTheirOwn'!B3,Pacific!B3:B7,Southeast!B3:B7,Southwest!B3:B7)",
                TestUtility.getFormulaCellValue(static1, 0, 1));

        List<String> divisionNames = Arrays.asList("Atlantic", "Central", "Southeast", "Northwest", "Pacific", "Southwest",
                "Empty", "Of Their Own", "DNE", "DNE-1");
        List<Integer> divisionTeams = Arrays.asList(5, 5, 5, 5, 5, 5, 0, 1, 0, 0);

        for (int i = 0; i < divisionNames.size(); i++)
        {
            Sheet division = workbook.getSheetAt(i + 1);
            String sheetName = divisionNames.get(i);
            String divisionName = sheetName;
            // Only the sheet name has "-1" appended due to duplicate sheet names.
            if (i == 9)
                divisionName = "DNE";
            int numTeams = divisionTeams.get(i);

            assertEquals("i: " + i, sheetName, division.getSheetName());
            Header header = division.getHeader();
            assertEquals("i: " + i, "Division: " + divisionName, header.getCenter());
            Footer footer = division.getFooter();
            assertEquals("i: " + i, "Division: " + divisionName, footer.getCenter());
            assertEquals("i: " + i, "Division: " + divisionName, TestUtility.getStringCellValue(division, 0, 0));
            if (numTeams > 1)
                assertEquals("i: " + i, "COUNTA(B3:B" + (numTeams + 2) + ")", TestUtility.getFormulaCellValue(division, numTeams + 2, 1));
            else if (numTeams == 1)
                assertEquals("i: " + i, "COUNTA(B3)", TestUtility.getFormulaCellValue(division, 3, 1));
            else
                assertEquals("i: " + i, "COUNTA($Z$1)", TestUtility.getFormulaCellValue(division, 2, 1));
            assertEquals("i: " + i, "n: " + i, TestUtility.getStringCellValue(division, numTeams + 3, 0));
            assertEquals("i: " + i, "s.index: " + i, TestUtility.getStringCellValue(division, numTeams + 4, 0));
            assertEquals("i: " + i, "s.first: " + (i == 0), TestUtility.getStringCellValue(division, numTeams + 5, 0));
            assertEquals("i: " + i, "s.last: " + (i == divisionNames.size() - 1), TestUtility.getStringCellValue(division, numTeams + 6, 0));
            assertEquals("i: " + i, "s.numIterations: " + divisionNames.size(), TestUtility.getStringCellValue(division, numTeams + 7, 0));
        }

        Sheet static2 = workbook.getSheetAt(11);
        assertEquals("COUNTA(Atlantic!B3:B7,Central!B3:B7,Northwest!B3:B7,'OfTheirOwn'!B3,Pacific!B3:B7,Southeast!B3:B7,Southwest!B3:B7)",
                TestUtility.getFormulaCellValue(static2, 0, 1));

        Sheet empty_1 = workbook.getSheetAt(12);
        assertEquals("empty-1", empty_1.getSheetName());
        assertEquals("COUNTA($Z$1)", TestUtility.getFormulaCellValue(empty_1, 2, 1));

        Sheet atlantic1 = workbook.getSheetAt(14);
        assertEquals("Atlantic-1", atlantic1.getSheetName());
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(atlantic1, 0, 0));
    }
}
