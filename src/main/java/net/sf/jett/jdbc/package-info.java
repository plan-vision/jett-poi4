/**
 * <p>Provides classes for executing queries via JDBC.</p>
 *
 * <p>The <code>JDBCExecutor</code> class, when an instance is exposed as a
 * bean, allows SQL queries to be run while specified in the template:</p>
 *
 * <code>&lt;jt:forEach items="${jdbc.execQuery('SELECT * FROM employee')}" var="employee"&gt;</code>
 *
 * <p>The <code>ResultSetRow</code> represents one row of data returned by a
 * <code>JDBCExecutor</code>.  It is not seen directly, but <code>JDBCExecutor</code>'s
 * <code>execQuery</code> method returns a <code>List</code> of <code>ResultSetRows</code>.</p>
 *
 * @author Randy Gettman
 * @since 0.6.0
 */
package net.sf.jett.jdbc;