package net.sf.jett.test;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.jdbc.JDBCExecutor;
import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the <code>JDBCExecutor</code> and
 * <code>ResultSetRow</code> classes.
 *
 * @author Randy Gettman
 * @since 0.6.0
 */
public class JDBCExecutorTest extends TestCase
{
    private static Connection theConnection;

    /**
     * Login to the HSQL DB prior to running any tests.
     * @throws ClassNotFoundException If the HSQL DB driver was not found.
     * @throws SQLException If there is an error connecting to the database.
     */
    @BeforeClass
    public static void setUpOnce() throws ClassNotFoundException, SQLException
    {
        Class.forName("org.hsqldb.jdbcDriver");
        theConnection = DriverManager.getConnection("jdbc:hsqldb:file:jett-db", "sa", "");

        try
        {
            update("DROP TABLE employee");
        }
        catch (SQLException ignored) {}
        update("CREATE TABLE employee (emp_id INTEGER, first_name VARCHAR(30), last_name VARCHAR(30), " +
                "salary INTEGER, title VARCHAR(30), manager VARCHAR(60), catch_phrase VARCHAR(100), is_a_manager VARCHAR(1))");
        update("INSERT INTO employee VALUES (1, 'Robert', 'Stack', 1000, 'Data Structures Programmer', null           , 'Push, Pop!'                       , 'Y')");
        update("INSERT INTO employee VALUES (2, 'Suzie',  'Queue',  900, 'Data Structures Programmer', 'Stack, Robert', 'Enqueue, Dequeue!'                , 'N')");
        update("INSERT INTO employee VALUES (3, 'Elmer',  'Fudd',   800, 'Cartoon Character'         , 'Bunny, Bugs'  , 'I''m hunting wabbits!  Huh-uh-uh!', 'N')");
        update("INSERT INTO employee VALUES (4, 'Bugs',   'Bunny', 1500, 'Cartoon Character'         , null           , 'Ah, what''s up Doc?'              , 'Y')");
    }

    /**
     * Runs the given SQL Statement via a <code>Statement</code> and a call to
     * <code>executeUpdate</code>.
     * @param sql The SQL Statement.
     * @throws SQLException If there is a problem executing the query.
     */
    private static void update(String sql) throws SQLException
    {
        Statement st = theConnection.createStatement();
        st.executeUpdate(sql);
        st.close();
    }

    /**
     * Close the connection to the HSQL DB after all tests are run.
     * @throws SQLException If there is a problem closing the connection.
     */
    @AfterClass
    public static void cleanUpOnce() throws SQLException
    {
        if (theConnection != null)
        {
            Statement st = theConnection.createStatement();
            st.execute("SHUTDOWN");
            theConnection.close();
        }
    }

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
        return "JDBCExecutor";
    }

    /**
     * Silence the transformer, which would complain when expressions evaluate to
     * <code>null</code>.  This is added so that <code>null</code>s can be
     * explicitly tested for in <code>JdbcExecutor</code>.
     * @param transformer The <code>ExcelTransformer</code>.
     * @since 0.9.0
     */
    @Override
    protected void setupTransformer(ExcelTransformer transformer)
    {
        transformer.setSilent(true);
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet query = workbook.getSheetAt(0);
        assertEquals("Robert", TestUtility.getStringCellValue(query, 1, 0));
        assertEquals("Queue", TestUtility.getStringCellValue(query, 2, 1));
        assertEquals(800, TestUtility.getNumericCellValue(query, 3, 2), Double.MIN_VALUE);
        assertEquals("Cartoon Character", TestUtility.getStringCellValue(query, 4, 3));
        assertTrue(TestUtility.isCellBlank(query, 1, 4));
        assertEquals("I'm hunting wabbits!  Huh-uh-uh!", TestUtility.getStringCellValue(query, 3, 5));
        assertTrue(TestUtility.isCellBlank(query, 4, 4));
        assertEquals("Y", TestUtility.getStringCellValue(query, 4, 6));

        Sheet prepared = workbook.getSheetAt(1);
        assertEquals("Cartoon Character", TestUtility.getStringCellValue(prepared, 0, 0));
        assertEquals("Cartoon Character", TestUtility.getStringCellValue(prepared, 2, 3));
        assertEquals("Cartoon Character", TestUtility.getStringCellValue(prepared, 3, 3));
        assertTrue(TestUtility.isCellBlank(prepared, 3, 4));
        assertEquals("Data Structures Programmer", TestUtility.getStringCellValue(prepared, 4, 0));
        assertEquals("Data Structures Programmer", TestUtility.getStringCellValue(prepared, 6, 3));
        assertTrue(TestUtility.isCellBlank(prepared, 6, 4));
        assertEquals("Data Structures Programmer", TestUtility.getStringCellValue(prepared, 7, 3));
        assertEquals("Nonexistent Title", TestUtility.getStringCellValue(prepared, 8, 0));
        assertTrue(TestUtility.isCellBlank(prepared, 10, 3));
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
        JDBCExecutor jdbc = new JDBCExecutor(theConnection);
        beans.put("jdbc", jdbc);

        List<String> titleSearches = Arrays.asList("Cartoon Character", "Data Structures Programmer", "Nonexistent Title");
        beans.put("titleSearches", titleSearches);
        return beans;
    }
}
