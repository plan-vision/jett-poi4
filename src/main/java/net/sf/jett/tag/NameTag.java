package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFName;

import net.sf.jett.exception.AttributeExpressionException;
import net.sf.jett.exception.TagParseException;
import net.sf.jett.formula.Formula;
import net.sf.jett.formula.CellRef;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.FormulaUtil;

/**
 * <p>A <code>NameTag</code> indicates an association between an Excel named
 * range and a JETT formula.  Instead of the JETT formula being transformed into
 * an Excel formula in the cell in which it's located, it will instead be
 * applied to the specified named range.  JETT does not verify that the
 * dynamically generated expression is a valid Excel Formula.  A
 * <code>NameTag</code> must be bodiless.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>name (required): <code>String</code></li>
 * <li>preferWorkbookScope (optional): <code>boolean</code></li>
 * <li>formula (required): <code>String</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public class NameTag extends BaseTag
{
    /**
     * A list of Excel built-in named range names.  These are hidden from
     * Excel's "Name Manager", but they cause a conflict when JETT attempts to
     * copy them from a source template sheet to a cloned template sheet.  They
     * appear to be covered by setting other sheet-related properties, e.g.
     * "Print_Titles" is covered by <code>setRepeatingRows</code> and
     * <code>setRepeatingColumns</code>.  This is a blacklist JETT uses to
     * bypass copying built-in names when cloning sheets.
     */
    public static final List<String> EXCEL_BUILT_IN_NAMES = Arrays.asList(
            // HSSF Built-in Names, found in org/apache/poi/hssf/record/NameRecord.java.
           /* 1 */ "Consolidate_Area", /* 2 */ "Auto_Open", /* 3 */ "Auto_Close", /* 4 */ "Database",
           /* 5 */ "Criteria", /* 6 */ "Print_Area", /* 7 */ "Print_Titles", /* 8 */ "Recorder",
           /* 9 */ "Data_Form", /* 10 */ "Auto_Activate", /* 11 */ "Auto_Deactivate", /* 12 */ "Sheet_Title",
           /* 13 */ "_FilterDatabase",
            // XSSF Built-in Names, found in the XSSFName class
            XSSFName.BUILTIN_CONSOLIDATE_AREA, XSSFName.BUILTIN_CRITERIA, XSSFName.BUILTIN_DATABASE,
            XSSFName.BUILTIN_EXTRACT, XSSFName.BUILTIN_FILTER_DB, XSSFName.BUILTIN_PRINT_AREA,
            XSSFName.BUILTIN_PRINT_TITLE, XSSFName.BUILTIN_SHEET_TITLE
    );

    /**
     * Attribute that specifies the name of the named range.
     */
    public static final String ATTR_NAME = "name";
    /**
     * Attribute that specifies whether to prefer a named range with workbook
     * scope over a same-named named range with sheet scope.  The default is
     * <code>false</code>, to prefer sheet scope over workbook scope.
     */
    public static final String ATTR_PREFER_WORKBOOK_SCOPE = "preferWorkbookScope";
    /**
     * Attribute that specifies the JETT formula to be applied to the named
     * range.  It must be enclosed in <code>$[</code> and <code>]</code> as
     * normal for JETT formulas.
     */
    public static final String ATTR_FORMULA = "formula";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_NAME, ATTR_FORMULA));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_PREFER_WORKBOOK_SCOPE));

    private Name myNamedRange;
    private String myJettFormula;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "name";
    }

    /**
     * Returns a <code>List</code> of required attribute names.
     * @return A <code>List</code> of required attribute names.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        List<String> reqAttrs = new ArrayList<>(super.getRequiredAttributes());
        reqAttrs.addAll(REQ_ATTRS);
        return reqAttrs;
    }

    /**
     * Returns a <code>List</code> of optional attribute names.
     * @return A <code>List</code> of optional attribute names.
     */
    @Override
    protected List<String> getOptionalAttributes()
    {
        List<String> optAttrs = new ArrayList<>(super.getOptionalAttributes());
        optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  The "name"
     * attribute must refer to an existing Excel named range name.  The
     * "formula" must be a JETT formula.
     */
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (!isBodiless())
            throw new TagParseException("Name tags must not have a body.  Name tag with body found" + getLocation());

        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Workbook workbook = sheet.getWorkbook();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();

        String name = AttributeUtil.evaluateStringNotNull(this, attributes.get(ATTR_NAME), beans, ATTR_NAME, "");

        boolean preferWorkbookScopeFirst = AttributeUtil.evaluateBoolean(this,
                attributes.get(ATTR_PREFER_WORKBOOK_SCOPE), beans, false);

        int numNamedRanges = workbook.getNumberOfNames();
        String sheetName = sheet.getSheetName();
        myNamedRange = null;
        for (int i = 0; i < numNamedRanges; i++)
        {
            Name namedRange = workbook.getNameAt(i);
            if (preferWorkbookScopeFirst)
            {
                if (namedRange.getSheetIndex() == -1 && namedRange.getNameName().equals(name))
                {
                    myNamedRange = namedRange;
                    break;
                }
                else if (sheetName.equals(namedRange.getSheetName()) && namedRange.getNameName().equals(name))
                {
                    myNamedRange = namedRange;
                    break;
                }
            }
            else
            {
                if (sheetName.equals(namedRange.getSheetName()) && namedRange.getNameName().equals(name))
                {
                    myNamedRange = namedRange;
                    break;
                }
                else if (namedRange.getSheetIndex() == -1 && namedRange.getNameName().equals(name))
                {
                    myNamedRange = namedRange;
                    break;
                }
            }
        }

        if (myNamedRange == null)
        {
            throw new AttributeExpressionException("NameTag: Unable to find named range with name \"" +
                    name + "\" in the workbook.  Reference found" + getLocation());
        }

        myJettFormula = AttributeUtil.evaluateStringNotNull(this, attributes.get(ATTR_FORMULA), beans, ATTR_FORMULA, "");
        if (!myJettFormula.startsWith(Formula.BEGIN_FORMULA) || !myJettFormula.endsWith(Formula.END_FORMULA))
        {
            throw new AttributeExpressionException("NameTag: Expected JETT formula of the form \"" +
                    Formula.BEGIN_FORMULA + "formula" + Formula.END_FORMULA + "\", got \"" + myJettFormula + "\"" +
                    getLocation());
        }
    }

    /**
     * <p>There should be a JETT formula in the formula map, keyed by the
     * location of this tag, referring to the JETT formula.  Replace the key
     * with a different key based on the named range instead.</p>
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        WorkbookContext workbookContext = getWorkbookContext();
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        String sheetName = sheet.getSheetName();
        String formulaMapKey = sheetName + "!" + myJettFormula;
        // Format: sheetName!name?scope
        // A "?" is chosen to distinguish this key from any other possible
        // formula key, because "?" isn't a legal character in a sheet name.
        String newKey = sheetName + "!" + myNamedRange.getNameName() + "?" + myNamedRange.getSheetName();

        // In the formula map, create another mapping of the named range key
        // referring to the same Formula.
        Map<String, Formula> formulaMap = workbookContext.getFormulaMap();
        Formula existingFormula = formulaMap.get(formulaMapKey);
        List<CellRef> existingFormulaCellRefs = existingFormula.getCellRefs();
        formulaMap.put(newKey, existingFormula);

        // In the cell ref map, make sure all mapped cell refs have the sheet
        // name referenced.
        Map<String, List<CellRef>> cellRefMap = workbookContext.getCellRefMap();
        for (CellRef formulaCellRef : existingFormulaCellRefs)
        {
            boolean isExplicit = formulaCellRef.getSheetName() != null;
            String cellKey = FormulaUtil.getCellKey(formulaCellRef, sheetName);
            if (isExplicit)
            {
                cellKey = FormulaUtil.EXPLICIT_REF_PREFIX + cellKey;
            }
            else
            {
                cellKey = FormulaUtil.IMPLICIT_REF_PREFIX + cellKey;
            }
            List<CellRef> cellRefs = cellRefMap.get(cellKey);
            for (int i = 0; i < cellRefs.size(); i++)
            {
                CellRef cellRef = cellRefs.get(i);
                String cellRefSheetName = cellRef.getSheetName();
                if (cellRefSheetName == null || "".equals(cellRefSheetName))
                {
                    CellRef newCellRef = new CellRef(sheetName, cellRef.getRow(), cellRef.getCol(),
                            cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                    // Replace any sheetless cell ref with a new cell ref that
                    // refers to the sheet on which this tag is located.
                    cellRefs.set(i, newCellRef);
                }
            }
        }

        // Clear the cell; there is no cell-visible result.
        clearBlock();
        return true;
    }
}
