package net.sf.jett.expression;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.util.CellReference;

/**
 * A <code>JettFuncs</code> object is an object that represents JETT utility
 * functionality in the JEXL world.
 *
 * @author Randy Gettman
 * @since 0.4.0
 */
public class JettFuncs
{
    private static final Integer[] CARDS = new Integer[52];
    private static int numCardsDealt = 0;

    /**
     * Takes 0-based row and column numbers (e.g. 1, 4), and generates an Excel cell
     * reference (e.g. "D2").
     * @param rowNum The 0-based row number.
     * @param colNum The 0-based column number.
     * @return A string representing an Excel cell reference.
     */
    public static String cellRef(int rowNum, int colNum)
    {
        return CellReference.convertNumToColString(colNum) + (rowNum + 1);
    }

    /**
     * Takes 0-based row and column numbers (e.g. 1, 4) and height and width
     * parameters (e.g. 2, 2), and generates an Excel cell
     * reference (e.g. "D2:E3").
     * @param rowNum The 0-based row number.
     * @param colNum The 0-based column number.
     * @param numRows The number of rows in the reference.
     * @param numCols The number of columns in the reference.
     * @return A string representing an Excel cell reference.
     */
    public static String cellRef(int rowNum, int colNum, int numRows, int numCols)
    {
        return CellReference.convertNumToColString(colNum) + (rowNum + 1) + ":" +
                CellReference.convertNumToColString(colNum + numCols - 1) + (rowNum + numRows);
    }

    /**
     * Picks a random card.
     * @return A random card.
     * @since 0.9.1
     */
    public static String pickACard()
    {
        if (numCardsDealt == 0)
        {
            for (int i = 0; i < CARDS.length; i++)
            {
                CARDS[i] = i;
            }
        }
        int index = numCardsDealt % CARDS.length;
        if (index == 0)
        {
            List<Integer> asList = Arrays.asList(CARDS);
            Collections.shuffle(asList);
        }
        int card = CARDS[index];
        int suit = card / 13;
        int rank = card % 13;

        numCardsDealt++;
        StringBuilder buf = new StringBuilder();
        switch (rank)
        {
        case 0:
            buf.append("Two");
            break;
        case 1:
            buf.append("Three");
            break;
        case 2:
            buf.append("Four");
            break;
        case 3:
            buf.append("Five");
            break;
        case 4:
            buf.append("Six");
            break;
        case 5:
            buf.append("Seven");
            break;
        case 6:
            buf.append("Eight");
            break;
        case 7:
            buf.append("Nine");
            break;
        case 8:
            buf.append("Ten");
            break;
        case 9:
            buf.append("Jack");
            break;
        case 10:
            buf.append("Queen");
            break;
        case 11:
            buf.append("King");
            break;
        case 12:
            buf.append("Ace");
            break;
        }
        buf.append(" of ");
        switch (suit)
        {
        case 0:
            buf.append("Clubs");
            break;
        case 1:
            buf.append("Diamonds");
            break;
        case 2:
            buf.append("Spades");
            break;
        case 3:
            buf.append("Hearts");
            break;
        }
        return buf.toString();
    }
}
