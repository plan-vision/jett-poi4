package net.sf.jett.test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the Formulas feature of JETT.
 *
 * @author Randy Gettman
 */
public class FormulaTest extends TestCase
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
     * @throws IOException If an I/O error occurs.
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
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "Formula";
    }

    /**
     * Call certain setup-related methods on the <code>ExcelTransformer</code>
     * before template sheet transformation.
     * @param transformer The <code>ExcelTransformer</code> that will transform
     *    the template worksheet(s).
     */
    @Override
    protected void setupTransformer(ExcelTransformer transformer)
    {
        transformer.setEvaluateFormulas(true);
        transformer.setForceRecalculationOnOpening(false);
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet formulaTest = workbook.getSheetAt(0);
        assertEquals("SUM(B3:B60)", TestUtility.getFormulaCellValue(formulaTest, 60, 1));
        assertEquals("SUM(C3:C60)", TestUtility.getFormulaCellValue(formulaTest, 60, 2));
        assertEquals("\"Counties:\"&COUNTA(E3:E60)", TestUtility.getFormulaCellValue(formulaTest, 60, 4));
        assertEquals("SUM(H3:H60)", TestUtility.getFormulaCellValue(formulaTest, 60, 7));
        assertEquals("SUM(I3:I60)", TestUtility.getFormulaCellValue(formulaTest, 60, 8));
        assertEquals("\"Counties:\"&COUNTA(K3:K60)", TestUtility.getFormulaCellValue(formulaTest, 60, 10));
        assertEquals("B61<>H61", TestUtility.getFormulaCellValue(formulaTest, 62, 2));
        assertEquals("TEXT(39300.625,\"[h]\")", TestUtility.getFormulaCellValue(formulaTest, 63, 2));

        // Test turning on formula evaluation.
        assertTrue(TestUtility.getBooleanCellValue(formulaTest, 62, 2));
        assertFalse(workbook.getForceFormulaRecalculation());

        for (int i = 1; i <= 6; i++)
        {
            Sheet division = workbook.getSheetAt(i);
            assertEquals("SUM(C3,D3)", TestUtility.getFormulaCellValue(division, 2, 4));
            assertEquals("SUM(C4,D4)", TestUtility.getFormulaCellValue(division, 3, 4));
            assertEquals("SUM(C5,D5)", TestUtility.getFormulaCellValue(division, 4, 4));
            assertEquals("SUM(C6,D6)", TestUtility.getFormulaCellValue(division, 5, 4));
            assertEquals("SUM(C7,D7)", TestUtility.getFormulaCellValue(division, 6, 4));
            assertEquals("COUNTA(B3:B7)", TestUtility.getFormulaCellValue(division, 7, 1));
            assertEquals("SUM(C3:C7)", TestUtility.getFormulaCellValue(division, 7, 2));
            assertEquals("SUM(D3:D7)", TestUtility.getFormulaCellValue(division, 7, 3));
            assertEquals("SUM(C3:C7,D3:D7)", TestUtility.getFormulaCellValue(division, 7, 4));
            assertEquals("SUM(C3:C7)/SUM(E3:E7)", TestUtility.getFormulaCellValue(division, 7, 5));
        }

        Sheet empty = workbook.getSheetAt(7);
        assertEquals("COUNTA($Z$1)", TestUtility.getFormulaCellValue(empty, 2, 1));
        assertEquals("SUM(0)", TestUtility.getFormulaCellValue(empty, 2, 2));
        assertEquals("SUM(0)", TestUtility.getFormulaCellValue(empty, 2, 3));
        assertEquals("SUM(0,0)", TestUtility.getFormulaCellValue(empty, 2, 4));
        assertEquals("SUM(0)/SUM(1)", TestUtility.getFormulaCellValue(empty, 2, 5));

        Sheet ofTheirOwn = workbook.getSheetAt(8);
        assertEquals("SUM(C3,D3)", TestUtility.getFormulaCellValue(ofTheirOwn, 2, 4));
        assertEquals("COUNTA(B3)", TestUtility.getFormulaCellValue(ofTheirOwn, 3, 1));
        assertEquals("SUM(C3)", TestUtility.getFormulaCellValue(ofTheirOwn, 3, 2));
        assertEquals("SUM(D3)", TestUtility.getFormulaCellValue(ofTheirOwn, 3, 3));
        assertEquals("SUM(C3,D3)", TestUtility.getFormulaCellValue(ofTheirOwn, 3, 4));
        assertEquals("SUM(C3)/SUM(E3)", TestUtility.getFormulaCellValue(ofTheirOwn, 3, 5));

        Sheet multiLevel = workbook.getSheetAt(9);
        assertEquals("COUNTA('FormulaTest'!$E$3:$E$60)", TestUtility.getFormulaCellValue(multiLevel, 0, 8));
        assertEquals("COUNTA('FormulaTest'!$K$3:$K$60)", TestUtility.getFormulaCellValue(multiLevel, 1, 8));

        assertEquals("SUM(C3,D3)", TestUtility.getFormulaCellValue(multiLevel, 2, 4));
        assertEquals("SUM(C4,D4)", TestUtility.getFormulaCellValue(multiLevel, 3, 4));
        assertEquals("SUM(C5,D5)", TestUtility.getFormulaCellValue(multiLevel, 4, 4));
        assertEquals("SUM(C6,D6)", TestUtility.getFormulaCellValue(multiLevel, 5, 4));
        assertEquals("SUM(C7,D7)", TestUtility.getFormulaCellValue(multiLevel, 6, 4));
        assertEquals("COUNTA(B3:B7)", TestUtility.getFormulaCellValue(multiLevel, 7, 1));
        assertEquals("SUM(C3:C7)", TestUtility.getFormulaCellValue(multiLevel, 7, 2));
        assertEquals("SUM(D3:D7)", TestUtility.getFormulaCellValue(multiLevel, 7, 3));
        assertEquals("SUM(E3:E7)", TestUtility.getFormulaCellValue(multiLevel, 7, 4));
        assertEquals("SUM(C3:C7)/SUM(E3:E7)", TestUtility.getFormulaCellValue(multiLevel, 7, 5));

        assertEquals("SUM(C11,D11)", TestUtility.getFormulaCellValue(multiLevel, 10, 4));
        assertEquals("SUM(C12,D12)", TestUtility.getFormulaCellValue(multiLevel, 11, 4));
        assertEquals("SUM(C13,D13)", TestUtility.getFormulaCellValue(multiLevel, 12, 4));
        assertEquals("SUM(C14,D14)", TestUtility.getFormulaCellValue(multiLevel, 13, 4));
        assertEquals("SUM(C15,D15)", TestUtility.getFormulaCellValue(multiLevel, 14, 4));
        assertEquals("COUNTA(B11:B15)", TestUtility.getFormulaCellValue(multiLevel, 15, 1));
        assertEquals("SUM(C11:C15)", TestUtility.getFormulaCellValue(multiLevel, 15, 2));
        assertEquals("SUM(D11:D15)", TestUtility.getFormulaCellValue(multiLevel, 15, 3));
        assertEquals("SUM(E11:E15)", TestUtility.getFormulaCellValue(multiLevel, 15, 4));
        assertEquals("SUM(C11:C15)/SUM(E11:E15)", TestUtility.getFormulaCellValue(multiLevel, 15, 5));

        assertEquals("SUM(C19,D19)", TestUtility.getFormulaCellValue(multiLevel, 18, 4));
        assertEquals("SUM(C20,D20)", TestUtility.getFormulaCellValue(multiLevel, 19, 4));
        assertEquals("SUM(C21,D21)", TestUtility.getFormulaCellValue(multiLevel, 20, 4));
        assertEquals("SUM(C22,D22)", TestUtility.getFormulaCellValue(multiLevel, 21, 4));
        assertEquals("SUM(C23,D23)", TestUtility.getFormulaCellValue(multiLevel, 22, 4));
        assertEquals("COUNTA(B19:B23)", TestUtility.getFormulaCellValue(multiLevel, 23, 1));
        assertEquals("SUM(C19:C23)", TestUtility.getFormulaCellValue(multiLevel, 23, 2));
        assertEquals("SUM(D19:D23)", TestUtility.getFormulaCellValue(multiLevel, 23, 3));
        assertEquals("SUM(E19:E23)", TestUtility.getFormulaCellValue(multiLevel, 23, 4));
        assertEquals("SUM(C19:C23)/SUM(E19:E23)", TestUtility.getFormulaCellValue(multiLevel, 23, 5));

        assertEquals("SUM(C27,D27)", TestUtility.getFormulaCellValue(multiLevel, 26, 4));
        assertEquals("SUM(C28,D28)", TestUtility.getFormulaCellValue(multiLevel, 27, 4));
        assertEquals("SUM(C29,D29)", TestUtility.getFormulaCellValue(multiLevel, 28, 4));
        assertEquals("SUM(C30,D30)", TestUtility.getFormulaCellValue(multiLevel, 29, 4));
        assertEquals("SUM(C31,D31)", TestUtility.getFormulaCellValue(multiLevel, 30, 4));
        assertEquals("COUNTA(B27:B31)", TestUtility.getFormulaCellValue(multiLevel, 31, 1));
        assertEquals("SUM(C27:C31)", TestUtility.getFormulaCellValue(multiLevel, 31, 2));
        assertEquals("SUM(D27:D31)", TestUtility.getFormulaCellValue(multiLevel, 31, 3));
        assertEquals("SUM(E27:E31)", TestUtility.getFormulaCellValue(multiLevel, 31, 4));
        assertEquals("SUM(C27:C31)/SUM(E27:E31)", TestUtility.getFormulaCellValue(multiLevel, 31, 5));

        assertEquals("SUM(C35,D35)", TestUtility.getFormulaCellValue(multiLevel, 34, 4));
        assertEquals("SUM(C36,D36)", TestUtility.getFormulaCellValue(multiLevel, 35, 4));
        assertEquals("SUM(C37,D37)", TestUtility.getFormulaCellValue(multiLevel, 36, 4));
        assertEquals("SUM(C38,D38)", TestUtility.getFormulaCellValue(multiLevel, 37, 4));
        assertEquals("SUM(C39,D39)", TestUtility.getFormulaCellValue(multiLevel, 38, 4));
        assertEquals("COUNTA(B35:B39)", TestUtility.getFormulaCellValue(multiLevel, 39, 1));
        assertEquals("SUM(C35:C39)", TestUtility.getFormulaCellValue(multiLevel, 39, 2));
        assertEquals("SUM(D35:D39)", TestUtility.getFormulaCellValue(multiLevel, 39, 3));
        assertEquals("SUM(E35:E39)", TestUtility.getFormulaCellValue(multiLevel, 39, 4));
        assertEquals("SUM(C35:C39)/SUM(E35:E39)", TestUtility.getFormulaCellValue(multiLevel, 39, 5));

        assertEquals("SUM(C43,D43)", TestUtility.getFormulaCellValue(multiLevel, 42, 4));
        assertEquals("SUM(C44,D44)", TestUtility.getFormulaCellValue(multiLevel, 43, 4));
        assertEquals("SUM(C45,D45)", TestUtility.getFormulaCellValue(multiLevel, 44, 4));
        assertEquals("SUM(C46,D46)", TestUtility.getFormulaCellValue(multiLevel, 45, 4));
        assertEquals("SUM(C47,D47)", TestUtility.getFormulaCellValue(multiLevel, 46, 4));
        assertEquals("COUNTA(B43:B47)", TestUtility.getFormulaCellValue(multiLevel, 47, 1));
        assertEquals("SUM(C43:C47)", TestUtility.getFormulaCellValue(multiLevel, 47, 2));
        assertEquals("SUM(D43:D47)", TestUtility.getFormulaCellValue(multiLevel, 47, 3));
        assertEquals("SUM(E43:E47)", TestUtility.getFormulaCellValue(multiLevel, 47, 4));
        assertEquals("SUM(C43:C47)/SUM(E43:E47)", TestUtility.getFormulaCellValue(multiLevel, 47, 5));

        assertEquals("COUNTA($Z$1)", TestUtility.getFormulaCellValue(multiLevel, 50, 1));
        assertEquals("SUM(0)", TestUtility.getFormulaCellValue(multiLevel, 50, 2));
        assertEquals("SUM(0)", TestUtility.getFormulaCellValue(multiLevel, 50, 3));
        assertEquals("SUM(0)", TestUtility.getFormulaCellValue(multiLevel, 50, 4));
        assertEquals("SUM(0)/SUM(1)", TestUtility.getFormulaCellValue(multiLevel, 50, 5));

        assertEquals("SUM(C54,D54)", TestUtility.getFormulaCellValue(multiLevel, 53, 4));
        assertEquals("COUNTA(B54)", TestUtility.getFormulaCellValue(multiLevel, 54, 1));
        assertEquals("SUM(C54)", TestUtility.getFormulaCellValue(multiLevel, 54, 2));
        assertEquals("SUM(D54)", TestUtility.getFormulaCellValue(multiLevel, 54, 3));
        assertEquals("SUM(E54)", TestUtility.getFormulaCellValue(multiLevel, 54, 4));
        assertEquals("SUM(C54)/SUM(E54)", TestUtility.getFormulaCellValue(multiLevel, 54, 5));

        assertEquals("COUNTA(B3:B7,B11:B15,B19:B23,B27:B31,B35:B39,B43:B47,B54)", TestUtility.getFormulaCellValue(multiLevel, 55, 1));
        assertEquals("SUM(C3:C7,C11:C15,C19:C23,C27:C31,C35:C39,C43:C47,C54)", TestUtility.getFormulaCellValue(multiLevel, 55, 2));
        assertEquals("SUM(D3:D7,D11:D15,D19:D23,D27:D31,D35:D39,D43:D47,D54)", TestUtility.getFormulaCellValue(multiLevel, 55, 3));
        assertEquals("SUM(E3:E7,E11:E15,E19:E23,E27:E31,E35:E39,E43:E47,E54)", TestUtility.getFormulaCellValue(multiLevel, 55, 4));
        assertEquals("SUM(C3:C7,C11:C15,C19:C23,C27:C31,C35:C39,C43:C47,C54)/SUM(E3:E7,E11:E15,E19:E23,E27:E31,E35:E39,E43:E47,E54)",
                TestUtility.getFormulaCellValue(multiLevel, 55, 5));

        Sheet copyRight = workbook.getSheetAt(10);
        assertEquals("SUM(A1+A2)", TestUtility.getFormulaCellValue(copyRight, 2, 0));
        assertEquals("SUM(B1+B2)", TestUtility.getFormulaCellValue(copyRight, 2, 1));
        assertEquals("SUM(C1+C2)", TestUtility.getFormulaCellValue(copyRight, 2, 2));
        assertEquals("SUM(D1+D2)", TestUtility.getFormulaCellValue(copyRight, 2, 3));
        assertEquals("SUM(E1+E2)", TestUtility.getFormulaCellValue(copyRight, 2, 4));
        assertEquals("SUM(F1+F2)", TestUtility.getFormulaCellValue(copyRight, 2, 5));
        assertEquals("SUM(G1+G2)", TestUtility.getFormulaCellValue(copyRight, 2, 6));
        assertEquals("SUM(H1+H2)", TestUtility.getFormulaCellValue(copyRight, 2, 7));
        assertEquals("SUM(I1+I2)", TestUtility.getFormulaCellValue(copyRight, 2, 8));
        assertEquals("SUM(J1+J2)", TestUtility.getFormulaCellValue(copyRight, 2, 9));

        Sheet replaceTest = workbook.getSheetAt(11);
        assertEquals("SUM(A1+A2)", TestUtility.getFormulaCellValue(replaceTest, 2, 0));
        assertEquals("SUM(A5+A6)", TestUtility.getFormulaCellValue(replaceTest, 6, 0));
        assertEquals("SUM(A9+A10)", TestUtility.getFormulaCellValue(replaceTest, 10, 0));
        assertEquals("SUM(A13+A14)", TestUtility.getFormulaCellValue(replaceTest, 14, 0));
        assertEquals("SUM(A17+A18)", TestUtility.getFormulaCellValue(replaceTest, 18, 0));
        assertEquals("SUM(A21+A22)", TestUtility.getFormulaCellValue(replaceTest, 22, 0));
        assertEquals("SUM(A25+A26)", TestUtility.getFormulaCellValue(replaceTest, 26, 0));
        assertEquals("SUM(A29+A30)", TestUtility.getFormulaCellValue(replaceTest, 30, 0));
        assertEquals("SUM(A33+A34)", TestUtility.getFormulaCellValue(replaceTest, 34, 0));

        Sheet outsideReference = workbook.getSheetAt(12);
        assertEquals("A1*B1", TestUtility.getFormulaCellValue(outsideReference, 0, 2));
        assertEquals("A1*B5", TestUtility.getFormulaCellValue(outsideReference, 4, 2));
        assertEquals("A1*B9", TestUtility.getFormulaCellValue(outsideReference, 8, 2));
        assertEquals("A1*B1*D1", TestUtility.getFormulaCellValue(outsideReference, 0, 4));
        assertEquals("A1*B1*D2", TestUtility.getFormulaCellValue(outsideReference, 1, 4));
        assertEquals("A1*B1*D3", TestUtility.getFormulaCellValue(outsideReference, 2, 4));
        assertEquals("A1*B1*D4", TestUtility.getFormulaCellValue(outsideReference, 3, 4));
        assertEquals("A1*B5*D5", TestUtility.getFormulaCellValue(outsideReference, 4, 4));
        assertEquals("A1*B5*D6", TestUtility.getFormulaCellValue(outsideReference, 5, 4));
        assertEquals("A1*B5*D7", TestUtility.getFormulaCellValue(outsideReference, 6, 4));
        assertEquals("A1*B5*D8", TestUtility.getFormulaCellValue(outsideReference, 7, 4));
        assertEquals("A1*B9*D9", TestUtility.getFormulaCellValue(outsideReference, 8, 4));
        assertEquals("A1*B9*D10", TestUtility.getFormulaCellValue(outsideReference, 9, 4));
        assertEquals("A1*B9*D11", TestUtility.getFormulaCellValue(outsideReference, 10, 4));
        assertEquals("A1*B9*D12", TestUtility.getFormulaCellValue(outsideReference, 11, 4));

        Sheet multiLevel2 = workbook.getSheetAt(13);
        assertEquals("SUM(B4)", TestUtility.getFormulaCellValue(multiLevel2, 4, 1));
        assertEquals("SUM(C4)", TestUtility.getFormulaCellValue(multiLevel2, 4, 2));
        assertEquals("SUM(D4)", TestUtility.getFormulaCellValue(multiLevel2, 4, 3));
        assertEquals("SUM(B7:B8)", TestUtility.getFormulaCellValue(multiLevel2, 8, 1));
        assertEquals("SUM(C7:C8)", TestUtility.getFormulaCellValue(multiLevel2, 8, 2));
        assertEquals("SUM(D7:D8)", TestUtility.getFormulaCellValue(multiLevel2, 8, 3));
        assertEquals("SUM(B5,B9)", TestUtility.getFormulaCellValue(multiLevel2, 10, 1));
        assertEquals("SUM(C5,C9)", TestUtility.getFormulaCellValue(multiLevel2, 10, 2));
        assertEquals("SUM(D5,D9)", TestUtility.getFormulaCellValue(multiLevel2, 10, 3));
        assertEquals("SUM(B13)", TestUtility.getFormulaCellValue(multiLevel2, 13, 1));
        assertEquals("SUM(C13)", TestUtility.getFormulaCellValue(multiLevel2, 13, 2));
        assertEquals("SUM(D13)", TestUtility.getFormulaCellValue(multiLevel2, 13, 3));
        assertEquals("SUM(B16)", TestUtility.getFormulaCellValue(multiLevel2, 16, 1));
        assertEquals("SUM(C16)", TestUtility.getFormulaCellValue(multiLevel2, 16, 2));
        assertEquals("SUM(D16)", TestUtility.getFormulaCellValue(multiLevel2, 16, 3));
        assertEquals("SUM(B14,B17)", TestUtility.getFormulaCellValue(multiLevel2, 18, 1));
        assertEquals("SUM(C14,C17)", TestUtility.getFormulaCellValue(multiLevel2, 18, 2));
        assertEquals("SUM(D14,D17)", TestUtility.getFormulaCellValue(multiLevel2, 18, 3));
        assertEquals("SUM(B21)", TestUtility.getFormulaCellValue(multiLevel2, 21, 1));
        assertEquals("SUM(C21)", TestUtility.getFormulaCellValue(multiLevel2, 21, 2));
        assertEquals("SUM(D21)", TestUtility.getFormulaCellValue(multiLevel2, 21, 3));
        assertEquals("SUM(B24)", TestUtility.getFormulaCellValue(multiLevel2, 24, 1));
        assertEquals("SUM(C24)", TestUtility.getFormulaCellValue(multiLevel2, 24, 2));
        assertEquals("SUM(D24)", TestUtility.getFormulaCellValue(multiLevel2, 24, 3));
        assertEquals("SUM(B22,B25)", TestUtility.getFormulaCellValue(multiLevel2, 26, 1));
        assertEquals("SUM(C22,C25)", TestUtility.getFormulaCellValue(multiLevel2, 26, 2));
        assertEquals("SUM(D22,D25)", TestUtility.getFormulaCellValue(multiLevel2, 26, 3));
        assertEquals("SUM(B29)", TestUtility.getFormulaCellValue(multiLevel2, 29, 1));
        assertEquals("SUM(C29)", TestUtility.getFormulaCellValue(multiLevel2, 29, 2));
        assertEquals("SUM(D29)", TestUtility.getFormulaCellValue(multiLevel2, 29, 3));
        assertEquals("SUM(B32)", TestUtility.getFormulaCellValue(multiLevel2, 32, 1));
        assertEquals("SUM(C32)", TestUtility.getFormulaCellValue(multiLevel2, 32, 2));
        assertEquals("SUM(D32)", TestUtility.getFormulaCellValue(multiLevel2, 32, 3));
        assertEquals("SUM(B30,B33)", TestUtility.getFormulaCellValue(multiLevel2, 34, 1));
        assertEquals("SUM(C30,C33)", TestUtility.getFormulaCellValue(multiLevel2, 34, 2));
        assertEquals("SUM(D30,D33)", TestUtility.getFormulaCellValue(multiLevel2, 34, 3));
        assertEquals("SUM(B11,B19,B27,B35)", TestUtility.getFormulaCellValue(multiLevel2, 36, 1));
        assertEquals("SUM(C11,C19,C27,C35)", TestUtility.getFormulaCellValue(multiLevel2, 36, 2));
        assertEquals("SUM(D11,D19,D27,D35)", TestUtility.getFormulaCellValue(multiLevel2, 36, 3));

        Sheet grid = workbook.getSheetAt(14);
        assertEquals("SUM(B2:D2)", TestUtility.getFormulaCellValue(grid, 1, 4));
        assertEquals("SUM(B3:D3)", TestUtility.getFormulaCellValue(grid, 2, 4));
        assertEquals("SUM(B4:D4)", TestUtility.getFormulaCellValue(grid, 3, 4));
        assertEquals("SUM(OFFSET(B2,0,0,3,1))", TestUtility.getFormulaCellValue(grid, 4, 1));
        assertEquals("SUM(OFFSET(B2,0,1,3,1))", TestUtility.getFormulaCellValue(grid, 4, 2));
        assertEquals("SUM(OFFSET(B2,0,2,3,1))", TestUtility.getFormulaCellValue(grid, 4, 3));
        assertEquals("SUM(E2:E4)", TestUtility.getFormulaCellValue(grid, 4, 4));

        Sheet tagParseInFormula = workbook.getSheetAt(15);
        assertEquals("MIN((N2>=$P$1)*(N2<$R$1))", TestUtility.getFormulaCellValue(tagParseInFormula, 1, 0));

        Sheet noSpaceAfterParen = workbook.getSheetAt(16);
        for (int r = 1; r <= 10; r++)
        {
            int er = r + 1;
            assertEquals("A" + er + "-(IF(B" + er + "=\"-\",0,B" + er + ")+C" + er + ")", TestUtility.getFormulaCellValue(noSpaceAfterParen, r, 3));
        }
    }

    /**
     * This test is a single map test.
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
        String[] templateSheetNameArray = new String[17];
        Arrays.fill(templateSheetNameArray, "Cloning");
        templateSheetNameArray[0] = "Formula Test";
        templateSheetNameArray[9] = "MultiLevel";
        templateSheetNameArray[10] = "Copy Right";
        templateSheetNameArray[11] = "ReplaceTest";
        templateSheetNameArray[12] = "Outside Reference";
        templateSheetNameArray[13] = "MultiLevel2";
        templateSheetNameArray[14] = "Grid";
        templateSheetNameArray[15] = "TagParseInFormula";
        templateSheetNameArray[16] = "NoSpaceAfterParen";
        return Arrays.asList(templateSheetNameArray);
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of result
     * sheet names.
     * @return A <code>List</code> of result sheet names.
     */
    @Override
    protected List<String> getListOfResultSheetNames()
    {
        return Arrays.asList("Formula Test", "Atlantic", "Central", "Southeast", "Northwest",
                "Pacific", "Southwest", "Empty", "Of Their Own", "MultiLevel", "Copy Right", "ReplaceTest",
                "Outside Reference", "MultiLevel2", "Grid", "TagParseInFormula", "NoSpaceAfterParen");
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
        List<Map<String, Object>> beansList = new ArrayList<>();
        beansList.add(TestUtility.getStateData());
        for (int i = 0; i < 8; i++)
            beansList.add(TestUtility.getSpecificDivisionData(i));
        beansList.add(TestUtility.getDivisionData());
        Map<String, Object> emptyBeans = new HashMap<>();
        // For "Copy Right" and "ReplaceTest".
        for (int f = 0; f < 2; f++)
            beansList.add(emptyBeans);
        // For "Outside Reference"
        Map<String, Object> outsideRefsBeans = new HashMap<>();
        outsideRefsBeans.put("two", 2);
        outsideRefsBeans.put("primes", Arrays.asList(3, 5, 7));
        outsideRefsBeans.put("morePrimes", Arrays.asList(11, 13, 17, 19));
        beansList.add(outsideRefsBeans);
        // For "Multilevel2"
        beansList.add(TestUtility.getWorkOrderData());
        // For "Grid"
        beansList.add(TestUtility.getRegionSalesData());
        // For "TagParseInFormula", "NoSpaceAfterParen"
        beansList.add(emptyBeans);
        beansList.add(emptyBeans);

        return beansList;
    }
}
