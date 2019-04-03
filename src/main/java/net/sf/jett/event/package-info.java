/**
 * <p>Provides support for event-based processing of template spreadsheets.
 * This includes <code>CellListeners</code> and <code>SheetListeners</code>,
 * which can be registered with the <code>ExcelTransformer</code>, and
 * <code>TagListeners</code> and <code>TagLoopListeners</code>, which can be
 * registered on any tags in the template spreadsheet.  Users can implement
 * these interfaces to supply custom processing.</p>
 *
 * @author Randy Gettman
 * @since 0.1.0
 */
package net.sf.jett.event;