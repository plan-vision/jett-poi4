package net.sf.jett.model;

/**
 * <p>A <code>PastEndAction</code> enumerated value specifies a possible action
 * when dealing with expressions that reference collection items beyond the end
 * of the iteration.  This comes up when a <code>MultiForEachTag</code> is
 * operating on collections of different sizes, and one collection has run out
 * of values before another collection.</p>
 *
 * <p>If such an expression were written in Java code, it would result in an
 * <code>IndexOutOfBoundsException</code>.  Each enumerated value specifies a
 * way of handling this condition.</p>
 *
 * @author Randy Gettman
 */
public enum PastEndAction
{
    /**
     * Specifies that any <code>Cell</code> containing an expression that
     * references a collection item beyond the end of the iteration should
     * result in the entire <code>Cell</code> being blanked out.
     */
    CLEAR_CELL,
    /**
     * Specifies that any <code>Cell</code> containing an expression that
     * references a collection item beyond the end of the iteration should
     * result in the entire <code>Cell</code> being removed, formatting and all.
     */
    REMOVE_CELL,
    /**
     * Specifies that any <code>Cell</code> containing an expression that
     * references a collection item beyond the end of the collection should
     * result only in those expressions containing a reference to the collection
     * item being replaced, e.g.
     * <code>${notBeyondCollection} and ${beyondCollection}</code> becomes
     * <code>NotBeyondValue and</code>, or <code>NotBeyondValue and -</code>,
     * depending on whether a specific replacement value is given.
     * @since 0.7.0
     */
    REPLACE_EXPR
}
