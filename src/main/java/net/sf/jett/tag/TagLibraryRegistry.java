package net.sf.jett.tag;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.parser.TagParser;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>TagLibraryRegistry</code> represents a registry for all
 * <code>TagLibraries</code> containing <code>Tags</code>.
 *
 * @author Randy Gettman
 */
public class TagLibraryRegistry
{
    private Map<String, TagLibrary> myRegistry;

    /**
     * Construct a <code>TagLibraryRegistry</code>.
     */
    public TagLibraryRegistry()
    {
        myRegistry = new HashMap<>();
    }

    /**
     * Add the given <code>TagLibrary</code> to the registry, associated with
     * the given namespace.
     * @param namespace The namespace.
     * @param library The <code>TagLibrary</code> to register.
     * @throws IllegalArgumentException If the namespace has already been
     *    registered.
     */
    public void registerTagLibrary(String namespace, TagLibrary library)
    {
        if (myRegistry.get(namespace) != null)
            throw new IllegalArgumentException("A tag library with namespace \"" +
                    namespace + "\" has already been registered.");
        myRegistry.put(namespace, library);
    }

    /**
     * Creates a <code>Tag</code>, looking in a specific namespace for a class
     * matching a specific tag name.  The given <code>TagParser</code> supplies
     * the namespace, the tag name, and the attributes.  If found, this creates
     * the <code>Tag</code> and gives it the given <code>TagContext</code>, else
     * it returns <code>null</code>.
     * @param parser A <code>TagParser</code> that has parsed tag text.
     * @param context The <code>TagContext</code>.
     * @param workbookContext The <code>WorkbookContext</code>.
     * @return A new <code>Tag</code>, or <code>null</code> if it couldn't be
     *    created.
     * @throws TagParseException If there was a problem instantiating the
     *    desired <code>Tag</code>.
     */
    public Tag createTag(TagParser parser, TagContext context, WorkbookContext workbookContext)
    {
        if (parser == null)
            return null;
        String namespace = parser.getNamespace();
        String tagName = parser.getTagName();
        Map<String, RichTextString> attributes = parser.getAttributes();
        if (namespace == null || tagName == null)
            return null;
        TagLibrary library = myRegistry.get(namespace);
        if (library == null)
            return null;
        Class<? extends Tag> tagClass = library.getTagMap().get(tagName);
        if (tagClass == null)
        {
            return null;
        }
        try
        {
            Tag tag = tagClass.newInstance();
            tag.setContext(context);
            tag.setWorkbookContext(workbookContext);
            tag.setAttributes(attributes);
            tag.setBodiless(parser.isBodiless());
            return tag;
        }
        catch (Exception e)
        {
            throw new TagParseException("Unable to create tag " + namespace + ":" + tagName +
                    SheetUtil.getCellLocation(parser.getCell()), e);
        }
    }
}

