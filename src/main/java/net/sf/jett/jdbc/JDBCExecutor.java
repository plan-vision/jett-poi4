package net.sf.jett.jdbc;

import java.io.BufferedReader;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URL;
import java.sql.Array;
import java.sql.Blob;
import java.sql.Clob;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.RowId;
import java.sql.SQLException;
import java.sql.SQLXML;
import java.sql.Statement;
import java.sql.Time;
import java.sql.Timestamp;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;

/**
 * A <code>JDBCExecutor</code> has the capability of running SQL statements via
 * JDBC over a supplied <code>Connection</code>.
 *
 * @author Randy Gettman
 * @since 0.6.0
 */
public class JDBCExecutor
{
    private static final Logger logger = LogManager.getLogger();

    private Connection myConnection;

    /**
     * Constructs a <code>JDBCExecutor</code> that will operate over the given
     * open <code>Connection</code> to a database.
     * @param connection An open <code>Connection</code>.
     */
    public JDBCExecutor(Connection connection)
   {
      myConnection = connection;
   }

    /**
     * Executes the given SQL statement using a <code>Statement</code> to obtain
     * a <code>List</code> of <code>ResultSetRows</code>.  Execution of the
     * <code>Statement</code> yields a <code>ResultSet</code>, which is
     * processed to create the <code>ResultSetRows</code>.
     * @param sql The SQL statement.
     * @return A <code>List</code> of <code>ResultSetRows</code>.
     * @throws SQLException If there was a problem executing the statement.
     */
    public List<ResultSetRow> execQuery(String sql) throws SQLException
    {
        try (Statement st = myConnection.createStatement(); ResultSet rs = st.executeQuery(sql))
        {
            return processResultSet(rs);
        }
    }

    /**
     * Executes the given SQL statement using a <code>PreparedStatement</code>
     * to obtain a <code>List</code> of <code>ResultSetRows</code>.  Execution
     * of the <code>PreparedStatement</code> yields a <code>ResultSet</code>,
     * which is processed to create the <code>ResultSetRows</code>.
     * @param sql The SQL statement.
     * @param bindVariableValues Optional bind variable values.  There must be
     *    exactly one of these for every <code>?</code> in the SQL query.
     * @return A <code>List</code> of <code>ResultSetRows</code>.
     * @throws SQLException If there was a problem executing the statement.
     */
    public List<ResultSetRow> execQuery(String sql, Object... bindVariableValues) throws SQLException
    {
        try (PreparedStatement ps = myConnection.prepareStatement(sql))
        {
            for (int i = 0; i < bindVariableValues.length; i++)
            {
                // Set bind variables here.
                // Try for most common first.
                // Convert to 1-based JDBC index.
                Object o = bindVariableValues[i];
                if (o instanceof String)
                    ps.setString(i + 1, (String) o);
                else if (o instanceof Integer)
                    ps.setInt(i + 1, (Integer) o);
                else if (o instanceof Double)
                    ps.setDouble(i + 1, (Double) o);
                else if (o instanceof Boolean)
                    ps.setBoolean(i + 1, (Boolean) o);
                else if (o instanceof Float)
                    ps.setFloat(i + 1, (Float) o);
                else if (o instanceof Long)
                    ps.setLong(i + 1, (Long) o);
                else if (o instanceof Date)
                    ps.setDate(i + 1, (Date) o);
                else if (o instanceof Time)
                    ps.setTime(i + 1, (Time) o);
                else if (o instanceof Timestamp)
                    ps.setTimestamp(i + 1, (Timestamp) o);
                else if (o instanceof BigDecimal)
                    ps.setBigDecimal(i + 1, (BigDecimal) o);
                else if (o instanceof Short)
                    ps.setShort(i + 1, (Short) o);
                else if (o instanceof Byte)
                    ps.setByte(i + 1, (Byte) o);
                else if (o instanceof byte[])
                    ps.setBytes(i + 1, (byte[]) o);
                else if (o instanceof Clob)
                    ps.setClob(i + 1, (Clob) o);
                else if (o instanceof Blob)
                    ps.setBlob(i + 1, (Blob) o);
                else if (o instanceof Array)
                    ps.setArray(i + 1, (Array) o);
                else if (o instanceof SQLXML)
                    ps.setSQLXML(i + 1, (SQLXML) o);
                else if (o instanceof RowId)
                    ps.setRowId(i + 1, (RowId) o);
                else if (o instanceof URL)
                    ps.setURL(i + 1, (URL) o);
                // Should cover NULL as well.
                else
                    ps.setObject(i + 1, o);
            }
            try (ResultSet rs = ps.executeQuery())
            {
                return processResultSet(rs);
            }
        }
    }

    /**
     * Processes the given <code>ResultSet</code>.  Reads all rows and all
     * content, placing values into <code>ResultSetRows</code>.
     * @param rs An unprocessed <code>ResultSet</code>.
     * @return A <code>List</code> of <code>ResultSetRows</code>.
     * @throws SQLException If there is a problem processing the result set.
     */
    private List<ResultSetRow> processResultSet(ResultSet rs) throws SQLException
    {
        ResultSetMetaData rsmd = rs.getMetaData();
        List<Integer> types = getTypes(rsmd);
        List<String> colNames = getColumnNames(rsmd);
        List<ResultSetRow> rows = new ArrayList<>();
        while (rs.next())
        {
            ResultSetRow row = new ResultSetRow();
            for (int i = 0; i < types.size(); i++)
            {
                // http://docs.oracle.com/javase/6/docs/technotes/guides/jdbc/getstart/mapping.html#996857
                // gives the mappings between JDBC types and Java data types.
                // Convert to 1-based JDBC index.
                String colName = colNames.get(i);
                logger.debug("pRS: i={}, colName={}", i, colNames.get(i));

                switch(types.get(i))
                {
                case Types.CHAR:
                case Types.VARCHAR:
                case Types.LONGVARCHAR:
                    row.set(colName, rs.getString(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.BINARY:
                case Types.VARBINARY:
                case Types.LONGVARBINARY:
                    row.set(colName, rs.getBytes(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.BIT:
                case Types.BOOLEAN:
                    row.set(colName, rs.getBoolean(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.TINYINT:
                case Types.SMALLINT:
                    row.set(colName, rs.getShort(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.INTEGER:
                    row.set(colName, rs.getInt(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.REAL:
                    row.set(colName, rs.getFloat(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.FLOAT:
                case Types.DOUBLE:
                    row.set(colName, rs.getDouble(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.DECIMAL:
                case Types.NUMERIC:
                    row.set(colName, rs.getBigDecimal(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.DATE:
                    row.set(colName, rs.getDate(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.TIME:
                    row.set(colName, rs.getTime(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                   break;
                case Types.TIMESTAMP:
                    row.set(colName, rs.getTimestamp(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.CLOB:
                {
                    Clob clob = rs.getClob(i + 1);
                    if (rs.wasNull())
                    {
                        row.set(colName, null);
                    }
                    else
                    {
                        BufferedReader r = new BufferedReader(clob.getCharacterStream());
                        StringBuffer buf = new StringBuffer();
                        String line;
                        try
                        {
                            while ((line = r.readLine()) != null)
                            {
                                buf.append(line);
                            }
                            row.set(colName, buf.toString());
                        }
                        catch (IOException e)
                        {
                            row.set(colName, e.getMessage());
                        }
                     }
                     break;
                }
                case Types.ARRAY:
                    row.set(colName, rs.getArray(i + 1).getArray());
                    if (rs.wasNull())
                        row.set(colName, null);
                    break;
                case Types.BLOB:
                case Types.JAVA_OBJECT:
                default:
                    row.set(colName, rs.getObject(i + 1));
                    if (rs.wasNull())
                        row.set(colName, null);
                }
            }
            rows.add(row);
        }

        return rows;
    }

    /**
     * Returns a <code>List</code> of all datatypes of all columns in the result set.
     * @param rsmd A <code>ResultSetMetaData</code>.
     * @return A <code>List</code> of <code>Integers</code> that represent the
     *    datatypes of the columns.
     * @throws SQLException If there is a problem accessing the metadata.
     * @see java.sql.Types
     */
    private List<Integer> getTypes(ResultSetMetaData rsmd) throws SQLException
    {
        int numCols = rsmd.getColumnCount();
        List<Integer> types = new ArrayList<>(numCols);
        for (int i = 0; i < numCols; i++)
        {
            // Convert to 1-based JDBC index.
            types.add(rsmd.getColumnType(i + 1));
        }
        return types;
    }

    /**
     * Returns a <code>List</code> of all column names in the result set.
     * @param rsmd A <code>ResultSetMetaData</code>.
     * @return A <code>List</code> of <code>Strings</code> that represent the
     *    column names.
     * @throws SQLException If there is a problem accessing the metadata.
    */
    private List<String> getColumnNames(ResultSetMetaData rsmd) throws SQLException
    {
        int numCols = rsmd.getColumnCount();
        List<String> colNames = new ArrayList<>(numCols);
        for (int i = 0; i < numCols; i++)
        {
            // Convert to 1-based JDBC index.
            colNames.add(rsmd.getColumnName(i + 1));
        }
        return colNames;
    }
}
