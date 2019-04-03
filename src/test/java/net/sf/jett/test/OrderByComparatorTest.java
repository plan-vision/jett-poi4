package net.sf.jett.test;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.exception.ParseException;
import net.sf.jett.test.model.Team;
import net.sf.jett.util.OrderByComparator;

/**
 * Test the <code>OrderByComparator</code>.
 *
 * @author Randy Gettman
 */
public class OrderByComparatorTest
{
    /**
     * Basic test with all four combinations.
     */
    @Test
    public void testOrderBy()
    {
        // Different percentages.
        Team team1 = new Team();
        team1.setCity("City1"); team1.setName("Name1"); team1.setWins(11); team1.setLosses(9);
        Team team2 = new Team();
        team2.setCity("City2"); team2.setName("Name2"); team2.setWins(9); team2.setLosses(11);
        // Same percentages, different number of wins.
        Team team3 = new Team();
        team3.setCity("City3"); team3.setName("Name3"); team3.setWins(10); team3.setLosses(5);
        Team team4 = new Team();
        team4.setCity("City4"); team4.setName("Name4"); team4.setWins(20); team4.setLosses(10);
        // Same percentages, same number of wins, different cities.
        Team team5 = new Team();
        team5.setCity("City5"); team5.setName("Name5"); team5.setWins(23); team5.setLosses(23);
        Team team6 = new Team();
        team6.setCity("City6"); team6.setName("Name6"); team6.setWins(23); team6.setLosses(23);
        Team team7 = new Team();
        team7.setCity(null);    team7.setName("Name7"); team7.setWins(23); team7.setLosses(23);
        // Same percentages, same number of wins, same city, different names.
        Team team8 = new Team();
        team8.setCity("City8"); team8.setName("Name8"); team8.setWins(30); team8.setLosses(10);
        Team team9 = new Team();
        team9.setCity("City8"); team9.setName(null);    team9.setWins(30); team9.setLosses(10);
        Team team10 = new Team();
        team10.setCity("City8"); team10.setName("Name10"); team10.setWins(30); team10.setLosses(10);

        List<Team> teams = Arrays.asList(team1, team2, team3, team4, team5, team6, team7, team8, team9, team10);
        List<String> orderByProps = Arrays.asList(
                "pct desc", "wins asc", "city asc nulls first", "name desc nulls last");
        OrderByComparator<Team> comp = new OrderByComparator<>(orderByProps);
        Collections.sort(teams, comp);

        List<Team> expected = Arrays.asList(team8, team10, team9, team3, team4, team1, team7, team5, team6, team2);
        for (int i = 0; i < 10; i++)
        {
            assertSame(expected.get(i), teams.get(i));
        }
    }

    /**
     * Test simple properties and defaults.
     */
    @Test
    public void testSimpleProperties()
    {
        List<String> properties = Arrays.asList("city", "name");
        OrderByComparator<Team> comp = new OrderByComparator<>(properties);

        List<String> expectedProps = Arrays.asList("city", "name");
        List<Integer> expectedOrderings = Arrays.asList(OrderByComparator.ORDER_ASC, OrderByComparator.ORDER_ASC);
        List<Integer> expectedNullOrderings = Arrays.asList(OrderByComparator.NULLS_LAST, OrderByComparator.NULLS_LAST);

        List<String> observedProps = comp.getProperties();
        List<Integer> observedOrderings = comp.getOrderings();
        List<Integer> observedNullOrderings = comp.getNullOrderings();

        assertEquals(2, observedProps.size());
        assertEquals(2, observedOrderings.size());
        assertEquals(2, observedNullOrderings.size());

        for (int i = 0; i < expectedProps.size(); i++)
        {
            assertEquals(expectedProps.get(i), observedProps.get(i));
            assertEquals(expectedOrderings.get(i), observedOrderings.get(i));
            assertEquals(expectedNullOrderings.get(i), observedNullOrderings.get(i));
        }
    }

    /**
     * Test simple properties and defaults.
     */
    @Test
    public void testAscDesc()
    {
        List<String> properties = Arrays.asList("city asc", "name desc");
        OrderByComparator<Team> comp = new OrderByComparator<>(properties);

        List<String> expectedProps = Arrays.asList("city", "name");
        List<Integer> expectedOrderings = Arrays.asList(OrderByComparator.ORDER_ASC, OrderByComparator.ORDER_DESC);
        List<Integer> expectedNullOrderings = Arrays.asList(OrderByComparator.NULLS_LAST, OrderByComparator.NULLS_FIRST);

        List<String> observedProps = comp.getProperties();
        List<Integer> observedOrderings = comp.getOrderings();
        List<Integer> observedNullOrderings = comp.getNullOrderings();

        assertEquals(2, observedProps.size());
        assertEquals(2, observedOrderings.size());
        assertEquals(2, observedNullOrderings.size());

        for (int i = 0; i < expectedProps.size(); i++)
        {
            assertEquals(expectedProps.get(i), observedProps.get(i));
            assertEquals(expectedOrderings.get(i), observedOrderings.get(i));
            assertEquals(expectedNullOrderings.get(i), observedNullOrderings.get(i));
        }
    }

    /**
     * Test simple null orderings.
     */
    @Test
    public void testNullOrderings()
    {
        List<String> properties = Arrays.asList("city nulls last", "name nulls first");
        OrderByComparator<Team> comp = new OrderByComparator<>(properties);

        List<String> expectedProps = Arrays.asList("city", "name");
        List<Integer> expectedOrderings = Arrays.asList(OrderByComparator.ORDER_ASC, OrderByComparator.ORDER_ASC);
        List<Integer> expectedNullOrderings = Arrays.asList(OrderByComparator.NULLS_LAST, OrderByComparator.NULLS_FIRST);

        List<String> observedProps = comp.getProperties();
        List<Integer> observedOrderings = comp.getOrderings();
        List<Integer> observedNullOrderings = comp.getNullOrderings();

        assertEquals(2, observedProps.size());
        assertEquals(2, observedOrderings.size());
        assertEquals(2, observedNullOrderings.size());

        for (int i = 0; i < expectedProps.size(); i++)
        {
            assertEquals(expectedProps.get(i), observedProps.get(i));
            assertEquals(expectedOrderings.get(i), observedOrderings.get(i));
            assertEquals(expectedNullOrderings.get(i), observedNullOrderings.get(i));
        }
    }

    /**
     * Test combinations.
     */
    @Test
    public void testCombinations()
    {
        List<String> properties = Arrays.asList(
                "city asc nulls last", "name asc nulls first", "wins desc nulls last", "pct desc nulls first");
        OrderByComparator<Team> comp = new OrderByComparator<>(properties);

        List<String> expectedProps = Arrays.asList("city", "name", "wins", "pct");
        List<Integer> expectedOrderings = Arrays.asList(
                OrderByComparator.ORDER_ASC, OrderByComparator.ORDER_ASC, OrderByComparator.ORDER_DESC, OrderByComparator.ORDER_DESC);
        List<Integer> expectedNullOrderings = Arrays.asList(
                OrderByComparator.NULLS_LAST, OrderByComparator.NULLS_FIRST, OrderByComparator.NULLS_LAST, OrderByComparator.NULLS_FIRST);

        List<String> observedProps = comp.getProperties();
        List<Integer> observedOrderings = comp.getOrderings();
        List<Integer> observedNullOrderings = comp.getNullOrderings();

        assertEquals(4, observedProps.size());
        assertEquals(4, observedOrderings.size());
        assertEquals(4, observedNullOrderings.size());

        for (int i = 0; i < expectedProps.size(); i++)
        {
            assertEquals(expectedProps.get(i), observedProps.get(i));
            assertEquals(expectedOrderings.get(i), observedOrderings.get(i));
            assertEquals(expectedNullOrderings.get(i), observedNullOrderings.get(i));
        }
    }

    /**
     * Error if bad ordering.
     */
    @Test(expected = ParseException.class)
    public void testBadOrdering()
    {
        List<String> properties = Arrays.asList("city backwards");
        new OrderByComparator<Team>(properties);
    }

    /**
     * Error if "nulls" only.
     */
    @Test(expected = ParseException.class)
    public void testNullsOnly()
    {
        List<String> properties = Arrays.asList("city nulls");
        new OrderByComparator<Team>(properties);
    }

    /**
     * Error if bad null ordering.
     */
    @Test(expected = ParseException.class)
    public void testBadNullOrdering()
    {
        List<String> properties = Arrays.asList("city nulls second");
        new OrderByComparator<Team>(properties);
    }

    /**
     * Error if ordering exists but only "nulls" after.
     */
    @Test(expected = ParseException.class)
    public void testOrderingAndNullsOnly()
    {
        List<String> properties = Arrays.asList("city asc nulls");
        new OrderByComparator<Team>(properties);
    }

    /**
     * Error if too many fields in a property.
     */
    @Test(expected = ParseException.class)
    public void testTooManyFields()
    {
        List<String> properties = Arrays.asList("city asc desc nulls first");
        new OrderByComparator<Team>(properties);
    }
}
