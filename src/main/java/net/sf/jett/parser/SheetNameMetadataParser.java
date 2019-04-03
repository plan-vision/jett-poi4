package net.sf.jett.parser;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * A <code>SheetNameMetadataParser</code> is a <code>MetadataParser</code>
 * that disallows certain metadata keys that do not make sense when present as
 * part of a sheet name.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class SheetNameMetadataParser extends MetadataParser
{
    /**
     * Abbreviation for {@link MetadataParser#VAR_NAME_INDEXVAR}.
     */
    public static final String ABBR_INDEX_VAR = "i";
    /**
     * Abbreviation for {@link MetadataParser#VAR_NAME_LIMIT}.
     */
    public static final String ABBR_LIMIT = "l";
    /**
     * Abbreviation for {@link MetadataParser#VAR_NAME_REPLACE_VALUE}.
     */
    public static final String ABBR_REPLACE_VALUE = "r";
    /**
     * Abbreviation for {@link MetadataParser#VAR_NAME_VAR_STATUS}.
     */
    public static final String ABBR_VAR_STATUS = "v";

    private static final List<String> RESTRICTED_KEYS = Arrays.asList(
            VAR_NAME_EXTRA_ROWS, VAR_NAME_LEFT, VAR_NAME_RIGHT, VAR_NAME_COPY_RIGHT,
            VAR_NAME_FIXED, VAR_NAME_PAST_END_ACTION, VAR_NAME_GROUP_DIR, VAR_NAME_COLLAPSE,
            VAR_NAME_ON_LOOP_PROCESSED, VAR_NAME_ON_PROCESSED
    );

    /**
     * Create a <code>SheetNameMetadataParser</code>.
     */
    public SheetNameMetadataParser()
    {
        super();
    }

    /**
     * Create a <code>SheetNameMetadataParser</code> object that will parse the given
     * metadata text.
     * @param metadataText The text of the metadata.
     */
    public SheetNameMetadataParser(String metadataText)
    {
        super(metadataText);
    }

    /**
     * Restricts many metadata keys that don't make sense in the context of a
     * sheet name.
     * @param metadataKey The metadata key.
     * @return <code>true</code> if it's restricted, else <code>false</code>.
     */
    @Override
    protected boolean isRestricted(String metadataKey)
    {
        return RESTRICTED_KEYS.contains(metadataKey);
    }

    /**
     * This parser will recognize abbreviations in the returned map.
     * @return A <code>Map</code> of abbreviations to metadata keys.
     */
    @Override
    protected Map<String, String> getAbbreviations()
    {
        Map<String, String> abbr = new HashMap<>();
        abbr.put(ABBR_INDEX_VAR, VAR_NAME_INDEXVAR);
        abbr.put(ABBR_LIMIT, VAR_NAME_LIMIT);
        abbr.put(ABBR_REPLACE_VALUE, VAR_NAME_REPLACE_VALUE);
        abbr.put(ABBR_VAR_STATUS, VAR_NAME_VAR_STATUS);
        return abbr;
    }
}
