package net.sf.jett.test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.Ignore;

import net.sf.jett.model.ExcelColor;
import net.sf.jett.tag.HyperlinkTag;
import net.sf.jett.test.model.County;
import net.sf.jett.test.model.Division;
import net.sf.jett.test.model.Element;
import net.sf.jett.test.model.Employee;
import net.sf.jett.test.model.HyperlinkData;
import net.sf.jett.test.model.Region;
import net.sf.jett.test.model.State;
import net.sf.jett.test.model.Team;
import net.sf.jett.test.model.WorkOrder;
import net.sf.jett.util.SheetUtil;

/**
 * This utility class supplies beans maps for possibly multiple tests.  It also
 * supplies convenience methods for accessing spreadsheet data for testing
 * purposes.
 *
 * @author Randy Gettman
 */
@Ignore
public class TestUtility
{
    /**
     * Gets a beans map with <code>State</code> beans "california" and "nevada".
     * Each bean in turn is composed of many <code>County</code> beans.
     *
     * @return A <code>Map</code> of <code>State</code> beans.
     */
    public static Map<String, Object> getStateData()
    {
        Map<String, Object> beans = new HashMap<>();
        beans.put("california", getCalifornia());
        beans.put("nevada", getNevada());

        return beans;
    }

    /**
     * Returns a specific state under the given bean name.
     * @param code A code (currently 0 = California, 1 = Nevada).
     * @param name The bean name to create.
     * @return A <code>Map</code> of the bean name to the bean value -- the
     *    <code>State</code>.
     * @since 0.8.0
     */
    public static Map<String, Object> getSpecificStateData(int code, String name)
    {
        Map<String, Object> beans = new HashMap<>();
        State state = null;
        switch(code)
        {
        case 0:
            state = getCalifornia();
            break;
        case 1:
            state = getNevada();
            break;
        }
        beans.put(name, state);
        return beans;
    }

    /**
     * Gets a beans map with "counties" referring to a <code>List</code> of all
     * counties in California and in Nevada.
     * @return A <code>Map</code> of a list of <code>County</code> beans.
     * @since 0.9.0
     */
    public static Map<String, Object> getCountyData()
    {
        Map<String, Object> beans = new HashMap<>();
        State california = getCalifornia();
        State nevada = getNevada();
        List<County> counties = new ArrayList<>();
        counties.addAll(california.getCounties());
        counties.addAll(nevada.getCounties());

        beans.put("counties", counties);
        return beans;
    }

    /**
     * Get California state data.
     * @return A <code>State</code>.
     * @since 0.8.0
     */
    private static State getCalifornia()
    {
        State california = new State();
        california.setName("California");
        california.addCounty(new County("Los Angeles", 10363850, 10515, 1850, "Los Angeles", "06037"));
        california.addCounty(new County("San Diego", 3146274, 10888, 1850, "San Diego", "06073"));
        california.addCounty(new County("Orange", 3121251, 2046, 1889, "Santa Ana", "06059"));
        california.addCounty(new County("Riverside", 2088322, 18669, 1893, "Riverside", "06065"));
        california.addCounty(new County("San Bernardino", 2055766, 51960, 1853, "San Bernardino", "06071"));
        california.addCounty(new County("Santa Clara", 1837075, 3344, 1850, "Santa Clara", "06085"));
        california.addCounty(new County("Alameda", 1543000, 1911, 1853, "Oakland", "06001"));
        california.addCounty(new County("Sacramento", 1424415, 2502, 1850, "Sacramento", "06067"));
        california.addCounty(new County("Contra Costa", 1051674, 1865, 1850, "Martinez", "06013"));
        california.addCounty(new County("Fresno", 931098, 15444, 1856, "Fresno", "06019"));
        california.addCounty(new County("San Francisco", 845559, 122, 1850, "San Francisco", "06075"));
        california.addCounty(new County("Ventura", 831587, 4781, 1872, "Ventura", "06111"));
        california.addCounty(new County("Kern", 817517, 21088, 1866, "Bakersfield", "06029"));
        california.addCounty(new County("San Mateo", 739469, 1163, 1856, "San Mateo", "06081"));
        california.addCounty(new County("San Joaquin", 685660, 3623, 1850, "Stockton", "06077"));
        california.addCounty(new County("Stanislaus", 525903, 3872, 1854, "Modesto", "06099"));
        california.addCounty(new County("Sonoma", 484470, 4082, 1850, "Santa Rosa", "06097"));
        california.addCounty(new County("Tulare", 435254, 12494, 1852, "Visalia", "06107"));
        california.addCounty(new County("Santa Barbara", 428655, 7091, 1850, "Santa Barbara", "06083"));
        california.addCounty(new County("Monterey", 428549, 8604, 1850, "Salinas", "06053"));
        california.addCounty(new County("Solano", 426757, 2145, 1850, "Fairfield", "06095"));
        california.addCounty(new County("Placer", 333401, 3893, 1851, "Auburn", "06061"));
        california.addCounty(new County("San Luis Obispo", 269337, 8557, 1850, "San Luis Obispo", "06079"));
        california.addCounty(new County("Santa Cruz", 266519, 1155, 1850, "Santa Cruz", "06087"));
        california.addCounty(new County("Marin", 257406, 1347, 1850, "San Rafael", "06041"));
        california.addCounty(new County("Merced", 255250, 4996, 1855, "Merced", "06047"));
        california.addCounty(new County("Butte", 220407, 4248, 1850, "Oroville", "06007"));
        california.addCounty(new County("Yolo", 199066, 2621, 1850, "Woodland", "06113"));
        california.addCounty(new County("Shasta", 182236, 9806, 1850, "Redding", "06089"));
        california.addCounty(new County("El Dorado", 179722, 4434, 1850, "Placerville", "06017"));
        california.addCounty(new County("Imperial", 176158, 10813, 1907, "El Centro", "06025"));
        california.addCounty(new County("Kings", 154434, 3600, 1893, "Hanford", "06031"));
        california.addCounty(new County("Madera", 150887, 5537, 1893, "Madera", "06039"));
        california.addCounty(new County("Napa", 136704, 1953, 1850, "Napa", "06055"));
        california.addCounty(new County("Humboldt", 132821, 9254, 1853, "Eureka", "06023"));
        california.addCounty(new County("Nevada", 99186, 2481, 1851, "Nevada City", "06057"));
        california.addCounty(new County("Sutter", 95878, 1562, 1850, "Sutter", "06101"));
        california.addCounty(new County("Mendocino", 90163, 9088, 1850, "Ukiah", "06045"));
        california.addCounty(new County("Yuba", 71929, 1632, 1850, "Yuba", "06115"));
        california.addCounty(new County("Lake", 64059, 3258, 1861, "Lakeport", "06033"));
        california.addCounty(new County("Tehama", 62419, 7643, 1856, "Red Bluff", "06103"));
        california.addCounty(new County("San Benito", 57784, 3597, 1874, "Hollister", "06069"));
        california.addCounty(new County("Tuolumne", 56799, 5791, 1850, "Tuolumne", "06109"));
        california.addCounty(new County("Calaveras", 46127, 2642, 1850, "San Andreas", "06009"));
        california.addCounty(new County("Siskiyou", 45971, 16283, 1852, "Yreka", "06093"));
        california.addCounty(new County("Amador", 37943, 1536, 1854, "Jackson", "06005"));
        california.addCounty(new County("Lassen", 35757, 11805, 1864, "Susanville", "06035"));
        california.addCounty(new County("Del Norte", 29419, 2611, 1850, "Crescent City", "06015"));
        california.addCounty(new County("Glenn", 29195, 3406, 1891, "Willows", "06021"));
        california.addCounty(new County("Colusa", 21910, 2981, 1850, "Colusa", "06011"));
        california.addCounty(new County("Plumas", 20917, 6615, 1854, "Quincy", "06063"));
        california.addCounty(new County("Mariposa", 18406, 3758, 1850, "Mariposa", "06043"));
        california.addCounty(new County("Inyo", 18152, 26397, 1866, "Independence", "06027"));
        california.addCounty(new County("Trinity", 13966, 8234, 1850, "Weaverville", "06105"));
        california.addCounty(new County("Mono", 13759, 7884, 1861, "Bridgeport", "06051"));
        california.addCounty(new County("Modoc", 9702, 10215, 1874, "Alturas", "06049"));
        california.addCounty(new County("Sierra", 3380, 2468, 1852, "Downieville", "06091"));
        california.addCounty(new County("Alpine", 1222, 1914, 1864, "Markleeville", "06003"));
        return california;
    }

    /**
     * Get Nevada state data.
     * @return A <code>State</code>.
     * @since 0.8.0
     */
    private static State getNevada()
    {

        State nevada = new State();
        nevada.setName("Nevada");
        nevada.addCounty(new County("Clark", 1375765, 20489, 1908, "Las Vegas", "32003"));
        nevada.addCounty(new County("Washoe", 339486, 16426, 1861, "Reno", "32031"));
        nevada.addCounty(new County("Carson City", 52457, 373, 1969, "Carson City", "32510"));
        nevada.addCounty(new County("Elko", 45291, 44501, 1869, "Elko", "32007"));
        nevada.addCounty(new County("Douglas", 41259, 1839, 1861, "Minden", "32005"));
        nevada.addCounty(new County("Lyon", 34501, 5164, 1861, "Yerington", "32019"));
        nevada.addCounty(new County("Nye", 32485, 47001, 1864, "Tonopah", "32023"));
        nevada.addCounty(new County("Churchill", 23982, 12766, 1861, "Fallon", "32001"));
        nevada.addCounty(new County("Humboldt", 16106, 24988, 1861, "Winnemucca", "32013"));
        nevada.addCounty(new County("White Pine", 9181, 22991, 1869, "Ely", "32033"));
        nevada.addCounty(new County("Pershing", 6693, 15563, 1919, "Lovelock", "32027"));
        nevada.addCounty(new County("Lander", 5794, 14229, 1861, "Battle Mountain", "32015"));
        nevada.addCounty(new County("Mineral", 5071, 9731, 1911, "Hawthorne", "32021"));
        nevada.addCounty(new County("Lincoln", 4165, 27545, 1866, "Pioche", "32017"));
        nevada.addCounty(new County("Storey", 3399, 684, 1861, "Virginia City", "32029"));
        nevada.addCounty(new County("Eureka", 1651, 10816, 1873, "Eureka", "32011"));
        nevada.addCounty(new County("Esmeralda", 971, 9295, 1861, "Goldfield", "32009"));
        return nevada;
    }

    /**
     * Returns some fictional <code>Counties</code> in a collection known as
     * "county", to test a specific part of implicit collections processing.
     * @return A <code>Map</code> of beans containing a <code>List</code> of
     *     dummy <code>County</code> objects.
     * @since 0.5.2
     */
    public static Map<String, Object> getFictionalCountyData()
    {
        Map<String, Object> beans = new HashMap<>();
        List<County> county = new ArrayList<>();
        county.add(new County("Carburetor", 150, 7089, 2006, "Radiator Springs", "99001"));
        county.add(new County("Cobblestone", 45678, 855, 1960, "Bedrock", "99003"));
        county.add(new County("Gotham", 10000000, 790, 1940, "Gotham", "99005"));
        county.add(new County("Greenbow", 123456, 10234, 1994, "Greenbow City", "99007"));
        county.add(new County("Hazzard", 1234, 4567, 1979, "Hazzard City", "99009"));
        county.add(new County("Heimlich", 34567, 10001, 1997, "Arlen", "99011"));
        county.add(new County("Hill", 1567890, 5034, 1985, "Hill Valley", "99013"));
        county.add(new County("Mayberry", 23456, 7501, 1960, "Mayberry", "99015"));
        county.add(new County("Maycomb", 13579, 489, 1960, "Maycomb", "99017"));
        county.add(new County("Metropolis", 10900000, 790, 1939, "Metropolis", "99019"));
        county.add(new County("Mist", 800, 8602, 1974, "Lake Wobegon", "99021"));
        county.add(new County("Springfield", 1357920, 1268, 1989, "Springfield", "99023"));
        // We call the bean "county" here, and there is a property "countySeat"
        // (method "getCountySeat") that should NOT be affected by implicit
        // collections processing.
        beans.put("county", county);

        return beans;
    }

    /**
     * Gets a beans map with a <code>List</code> of division beans, most of
     * which contain <code>Team</code> beans.  The name is "divisionsList".
     * @return A <code>Map</code> of beans containing a <code>List</code> of
     *    all <code>Divisions</code>.
     */
    public static Map<String, Object> getDivisionData()
    {
        Map<String, Object> beans = new HashMap<>();
        List<Division> divisionsList = new ArrayList<>();

        divisionsList.add(getAtlanticDivision());
        divisionsList.add(getCentralDivision());
        divisionsList.add(getSoutheastDivision());
        divisionsList.add(getNorthwestDivision());
        divisionsList.add(getPacificDivision());
        divisionsList.add(getSouthwestDivision());
        divisionsList.add(getEmptyDivision());
        divisionsList.add(getOfTheirOwnDivision());

        beans.put("divisionsList", divisionsList);

        return beans;
    }

    /**
     * Get a beans map with only one <code>Division</code>, determined by the
     * <code>code</code> argument.  The name is "division".
     * @param code Determines with division, 0-7.
     * @return A <code>Map</code> of beans, containing a specific
     *    <code>Division</code>.
     */
    public static Map<String, Object> getSpecificDivisionData(int code)
    {
        return getSpecificDivisionData(code, "division");
    }

    /**
     * Get a beans map with only one <code>Division</code>, determined by the
     * <code>code</code> argument.  It is keyed by the given name.
     * @param code Determines which division, 0-7.
     * @param name This becomes the bean name of the <code>Division</code>.
     * @return A <code>Map</code> of beans, containing a specific
     *    <code>Division</code> with the given name.
     */
    public static Map<String, Object> getSpecificDivisionData(int code, String name)
    {
        Map<String, Object> beans = new HashMap<>();
        Division division = null;
        switch(code)
        {
        case 0:
            division = getAtlanticDivision();
            break;
        case 1:
            division = getCentralDivision();
            break;
        case 2:
            division = getSoutheastDivision();
            break;
        case 3:
            division = getNorthwestDivision();
            break;
        case 4:
            division = getPacificDivision();
            break;
        case 5:
            division = getSouthwestDivision();
            break;
        case 6:
            division = getEmptyDivision();
            break;
        case 7:
            division = getOfTheirOwnDivision();
            break;
        }
        beans.put(name, division);
        return beans;
    }

    /**
     * Gets a beans map with a <code>List</code> of 100 dummy division beans,
     * each containing 10 <code>Team</code> beans.  The name is "divisionsList".
     * @return A <code>Map</code> of beans containing a <code>List</code> of
     *    100 dummy <code>Divisions</code>.
     * @since 0.8.0
     */
    public static Map<String, Object> getDummyDivisionsData()
    {
        Map<String, Object> beans = new HashMap<>();
        List<Division> divisionsList = new ArrayList<>(100);
        for (int d = 0; d < 100; d++)
            divisionsList.add(getDummyDivision(d));

        beans.put("divisionsList", divisionsList);

        return beans;
    }

    /**
     * Get a dummy <code>Division</code>, with 10 dummy teams.
     * @param d Determines which dummy "division".
     * @return A dummy <code>Division</code> with 10 dummy teams.
     * @since 0.8.0
     */
    public static Division getDummyDivision(int d)
    {
        Division div = new Division();
        div.setName("Division " + (d + 1));
        for (int t = 0; t < 10; t++)
        {
            Team team = new Team(div);
            team.setCity("City " + (d * 10 + t + 1));
            team.setName("The Team " + (d * 10 + t + 1) + "ers");
            team.setWins((int) (Math.random() * 162));
            team.setLosses(162 - team.getWins());
            div.addTeam(team);
        }
        return div;
    }

    /**
     * Get a beans map with a <code>List</code> of <code>Teams</code> from all
     * <code>Divisions</code>.  The name is "teams".
     * @return A <code>Map</code> of beans, containing the <code>List</code> of
     *    <code>Teams</code>.
     * @since 0.3.0
     */
    public static Map<String, Object> getTeamsData()
    {
        Map<String, Object> beans = new HashMap<>();
        List<Team> teams = new ArrayList<>();

        teams.addAll(getAtlanticDivision().getTeams());
        teams.addAll(getCentralDivision().getTeams());
        teams.addAll(getSoutheastDivision().getTeams());
        teams.addAll(getNorthwestDivision().getTeams());
        teams.addAll(getPacificDivision().getTeams());
        teams.addAll(getSouthwestDivision().getTeams());
        teams.addAll(getEmptyDivision().getTeams());
        teams.addAll(getOfTheirOwnDivision().getTeams());

        beans.put("teams", teams);
        return beans;
    }

    /**
     * Return Atlantic Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getAtlanticDivision()
    {
        Division atlantic = new Division();
        atlantic.setName("Atlantic");
        Team boston = new Team(atlantic);
        boston.setCity("Boston"); boston.setName("Celtics"); boston.setWins(51); boston.setLosses(21);
        atlantic.addTeam(boston);
        Team phila = new Team(atlantic);
        phila.setCity("Philadelphia"); phila.setName("76ers"); phila.setWins(37); phila.setLosses(36);
        atlantic.addTeam(phila);
        Team newYork = new Team(atlantic);
        newYork.setCity("New York"); newYork.setName("Knicks"); newYork.setWins(35); newYork.setLosses(38);
        atlantic.addTeam(newYork);
        Team newJersey = new Team(atlantic);
        newJersey.setCity("New Jersey"); newJersey.setName("Nets"); newJersey.setWins(23); newJersey.setLosses(49);
        atlantic.addTeam(newJersey);
        Team toronto = new Team(atlantic);
        toronto.setCity("Toronto"); toronto.setName("Raptors"); toronto.setWins(20); toronto.setLosses(53);
        atlantic.addTeam(toronto);
        return atlantic;
    }

    /**
     * Return Central Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getCentralDivision()
    {
        Division central = new Division();
        central.setName("Central");
        Team chicago = new Team(central);
        chicago.setCity("Chicago"); chicago.setName("Bulls"); chicago.setWins(53); chicago.setLosses(19);
        central.addTeam(chicago);
        Team indiana = new Team(central);
        indiana.setCity("Indiana"); indiana.setName("Pacers"); indiana.setWins(32); indiana.setLosses(42);
        central.addTeam(indiana);
        Team milwaukee = new Team(central);
        milwaukee.setCity("Milwaukee"); milwaukee.setName("Bucks"); milwaukee.setWins(29); milwaukee.setLosses(43);
        central.addTeam(milwaukee);
        Team detroit = new Team(central);
        detroit.setCity("Detroit"); detroit.setName("Pistons"); detroit.setWins(26); detroit.setLosses(47);
        central.addTeam(detroit);
        Team cleveland = new Team(central);
        cleveland.setCity("Cleveland"); cleveland.setName("Cavaliers"); cleveland.setWins(14); cleveland.setLosses(58);
        central.addTeam(cleveland);
        return central;
    }

    /**
     * Return Southeast Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getSoutheastDivision()
    {
        Division southeast = new Division();
        southeast.setName("Southeast");
        Team miami = new Team(southeast);
        miami.setCity("Miami"); miami.setName("Heat"); miami.setWins(51); miami.setLosses(22);
        southeast.addTeam(miami);
        Team orlando = new Team(southeast);
        orlando.setCity("Orlando"); orlando.setName("Magic"); orlando.setWins(47); orlando.setLosses(26);
        southeast.addTeam(orlando);
        Team atlanta = new Team(southeast);
        atlanta.setCity("Atlanta"); atlanta.setName("Hawks"); atlanta.setWins(42); atlanta.setLosses(32);
        southeast.addTeam(atlanta);
        Team charlotte = new Team(southeast);
        charlotte.setCity("Charlotte"); charlotte.setName("Bobcats"); charlotte.setWins(30); charlotte.setLosses(42);
        southeast.addTeam(charlotte);
        Team wash = new Team(southeast);
        wash.setCity("Washington"); wash.setName("Wizards"); wash.setWins(17); wash.setLosses(55);
        southeast.addTeam(wash);
        return southeast;
    }

    /**
     * Return Northwest Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getNorthwestDivision()
    {
        Division northwest = new Division();
        northwest.setName("Northwest");
        Team okCity = new Team(northwest);
        okCity.setCity("Oklahoma City"); okCity.setName("Thunder"); okCity.setWins(48); okCity.setLosses(24);
        northwest.addTeam(okCity);
        Team denver = new Team(northwest);
        denver.setCity("Denver"); denver.setName("Nuggets"); denver.setWins(44); denver.setLosses(29);
        northwest.addTeam(denver);
        Team portland = new Team(northwest);
        portland.setCity("Portland"); portland.setName("Trailblazers"); portland.setWins(42); portland.setLosses(31);
        northwest.addTeam(portland);
        Team utah = new Team(northwest);
        utah.setCity("Utah"); utah.setName("Jazz"); utah.setWins(36); utah.setLosses(38);
        northwest.addTeam(utah);
        Team minnesota = new Team(northwest);
        minnesota.setCity("Minnesota"); minnesota.setName("Timberwolves"); minnesota.setWins(17); minnesota.setLosses(57);
        northwest.addTeam(minnesota);
        return northwest;
    }

    /**
     * Return Pacific Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getPacificDivision()
    {
        Division pacific = new Division();
        pacific.setName("Pacific");
        Team lal = new Team(pacific);
        lal.setCity("Los Angeles"); lal.setName("Lakers"); lal.setWins(53); lal.setLosses(20);
        pacific.addTeam(lal);
        Team phoenix = new Team(pacific);
        phoenix.setCity("Phoenix"); phoenix.setName("Suns"); phoenix.setWins(36); phoenix.setLosses(36);
        pacific.addTeam(phoenix);
        Team gState = new Team(pacific);
        gState.setCity("Golden State"); gState.setName("Warriors"); gState.setWins(32); gState.setLosses(42);
        pacific.addTeam(gState);
        Team lac = new Team(pacific);
        lac.setCity("Los Angeles"); lac.setName("Clippers"); lac.setWins(29); lac.setLosses(45);
        pacific.addTeam(lac);
        Team sacramento = new Team(pacific);
        sacramento.setCity("Sacramento"); sacramento.setName("Kings"); sacramento.setWins(20); sacramento.setLosses(52);
        pacific.addTeam(sacramento);
        return pacific;
    }

    /**
     * Return Southwest Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getSouthwestDivision()
    {
        Division southwest = new Division();
        southwest.setName("Southwest");
        Team sanAnt = new Team(southwest);
        sanAnt.setCity("San Antonio"); sanAnt.setName("Spurs"); sanAnt.setWins(57); sanAnt.setLosses(16);
        southwest.addTeam(sanAnt);
        Team dallas = new Team(southwest);
        dallas.setCity("Dallas"); dallas.setName("Mavericks"); dallas.setWins(52); dallas.setLosses(21);
        southwest.addTeam(dallas);
        Team newOrl = new Team(southwest);
        newOrl.setCity("New Orleans"); newOrl.setName("Hornets"); newOrl.setWins(42); newOrl.setLosses(32);
        southwest.addTeam(newOrl);
        Team memphis = new Team(southwest);
        memphis.setCity("Memphis"); memphis.setName("Grizzlies"); memphis.setWins(41); memphis.setLosses(33);
        southwest.addTeam(memphis);
        Team houston = new Team(southwest);
        houston.setCity("Houston"); houston.setName("Rockets"); houston.setWins(38); houston.setLosses(35);
        southwest.addTeam(houston);
        return southwest;
    }

    /**
     * Return Empty Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getEmptyDivision()
    {
        Division empty = new Division();
        empty.setName("Empty");
        return empty;
    }

    /**
     * Return Of Their Own Division statistics.
     * @return A <code>Division</code>.
     */
    private static Division getOfTheirOwnDivision()
    {
        Division ofTheirOwn = new Division();
        ofTheirOwn.setName("Of Their Own");
        Team harlem = new Team(ofTheirOwn);
        harlem.setCity("Harlem"); harlem.setName("Globetrotters"); harlem.setWins(21227); harlem.setLosses(341);
        ofTheirOwn.addTeam(harlem);
        return ofTheirOwn;
    }

    /**
     * Gets a beans map with <code>HyperlinkData</code> data, exposed as
     * "hyperlinks".
     *
     * @return A <code>Map</code> of <code>HyperlinkData</code> beans.
     */
    public static Map<String, Object> getHyperlinkData()
    {
        Map<String, Object> beans = new HashMap<>();
        List<HyperlinkData> hyperlinks = new ArrayList<>();
        hyperlinks.add(new HyperlinkData(
                HyperlinkTag.TYPE_URL, "http://jett.sourceforge.net", "JETT on SourceForge"));
        hyperlinks.add(new HyperlinkData(
                HyperlinkTag.TYPE_EMAIL, "mailto:jett-users@lists.sourceforge.net", "Email jett-users"));
        hyperlinks.add(new HyperlinkData(
                HyperlinkTag.TYPE_FILE, "../templates/HyperlinkTagTemplate.xlsx", "Template For This Test (.xlsx)"));
        hyperlinks.add(new HyperlinkData(
                HyperlinkTag.TYPE_DOC, "'Target Sheet'!B3", "Intra-spreadsheet Link"));
        beans.put("hyperlinks", hyperlinks);
        return beans;
    }

    /**
     * Gets a beans map with <code>Employee</code> data, exposed as "employees".
     * @return A <code>Map</code> of <code>Employee</code> beans.
     * @since 0.3.0
     */
    public static Map<String, Object> getEmployeeData()
    {
        Map<String, Object> beans = new HashMap<>();
        beans.put("employees", getEmployees());
        return beans;
    }

    /**
     * Returns a <code>List</code> of <code>Employees</code>.
     * @return A <code>List</code> of <code>Employees</code>.
     * @since 0.6.0
     */
    public static List<Employee> getEmployees()
    {
        Employee robert = new Employee();
        robert.setFirstName("Robert");
        robert.setLastName("Stack");
        robert.setSalary(1000);
        robert.setTitle("Data Structures Programmer");
        robert.setAManager(true);
        Employee bugs = new Employee();
        bugs.setFirstName("Bugs");
        bugs.setLastName("Bunny");
        bugs.setSalary(1500);
        bugs.setCatchPhrase("Ah, what's up Doc?");
        bugs.setTitle("Cartoon Character");
        bugs.setAManager(true);
        Employee suzie = new Employee();
        suzie.setFirstName("Suzie");
        suzie.setLastName("Queue");
        suzie.setSalary(900);
        suzie.setManager(robert);
        suzie.setTitle("Data Structures Programmer");
        Employee elmer = new Employee();
        elmer.setFirstName("Elmer");
        elmer.setLastName("Fudd");
        elmer.setSalary(800);
        elmer.setManager(bugs);
        elmer.setCatchPhrase("I'm hunting wabbits!  Huh-uh-uh!");
        elmer.setTitle("Cartoon Character");
        return Arrays.asList(robert, suzie, elmer, bugs);
    }

    /**
     * Gets a beans map with <code>Element</code> data, exposed as "elements".
     * @return A <code>Map</code> of <code>Element</code> beans.
     * @since 0.9.1
     */
    public static Map<String, Object> getElementData()
    {
        Map<String, Object> beans = new HashMap<>();
        beans.put("elements", getElements());
        return beans;
    }

    /**
     * Returns a <code>List</code> of <code>Elements</code>.
     * @return A <code>List</code> of <code>Elements</code>.
     * @since 0.9.1
     */
    public static List<Element> getElements()
    {
        List<Element> elements = new ArrayList<>();
        List<Element> subElements1 = new ArrayList<>();
        List<Element> subElements2 = new ArrayList<>();
        subElements1.add(new Element("subTitle1", null));
        subElements1.add(new Element("subTitle2", null));
        subElements1.add(new Element("subTitle3", null));
        subElements2.add(new Element("subTitle4", null));
        elements.add(new Element("TITLE1", subElements1));
        elements.add(new Element("TITLE2", subElements2));
        return elements;
    }

    /**
     * Gets a beans map with <code>WorkOrder</code> data, exposed as "workOrders".
     * @return A <code>Map</code> of <code>WorkOrder</code> beans.
     * @since 0.10.0
     */
    public static Map<String, Object> getWorkOrderData()
    {
        Map<String, Object> beans = new HashMap<>();
        beans.put("workOrders", getWorkOrders());
        return beans;
    }

    /**
     * Returns a <code>List</code> of <code>WorkOrders</code>.
     * @return A <code>List</code> of <code>WorkOrders</code>.
     * @since 0.10.0
     */
    public static List<WorkOrder> getWorkOrders()
    {
        List<WorkOrder> workOrders = new ArrayList<>();
        // Must have at least 2 work orders with the same department, location,
        // and installation for this particular test to test its target issue
        // properly.
        workOrders.add(new WorkOrder("Hardware", "6th floor", "Linux Computers", "2015-11-03", 7800, 25500));
        workOrders.add(new WorkOrder("Hardware", "6th floor", "Windows Computers", "2015-11-02", 5500, 17500));
        workOrders.add(new WorkOrder("Hardware", "6th floor", "Windows Computers", "2015-11-03", 6000, 20000));
        workOrders.add(new WorkOrder("Hardware", "7th floor", "Linux Computers", "2015-11-04", 8200, 27400));
        workOrders.add(new WorkOrder("Hardware", "7th floor", "Windows Computers", "2015-11-04", 11200, 36500));
        workOrders.add(new WorkOrder("Software", "6th floor", "Open Office", "2015-11-03", 250, 0));
        workOrders.add(new WorkOrder("Software", "6th floor", "MS Office", "2015-11-03", 1000, 10000));
        workOrders.add(new WorkOrder("Software", "7th floor", "Open Office", "2015-11-04", 300, 0));
        workOrders.add(new WorkOrder("Software", "7th floor", "MS Office", "2015-11-04", 900, 10000));
        return workOrders;
    }

    /**
     * Gets a beans map with <code>Region</code> data, exposed as "regions".
     * Also exposes "dates" as the list of dates.
     * @return A <code>Map</code> of <code>Region</code> beans and date strings.
     * @since 0.10.0
     */
    public static Map<String, Object> getRegionSalesData()
    {
        Map<String, Object> beans = new HashMap<>();
        beans.put("regions", getRegions());
        beans.put("dates", Arrays.asList("1/1/2015", "2/1/2015", "3/1/2015"));
        return beans;
    }

    /**
     * Returns a <code>List</code> of <code>Regions</code>.
     * @return A <code>List</code> of <code>Regions</code>.
     * @since 0.10.0
     */
    public static List<Region> getRegions()
    {
        List<Region> regions = new ArrayList<>();
        regions.add(new Region("USA", Arrays.asList(50, 51, 52)));
        regions.add(new Region("ASIA", Arrays.asList(25, 26, 27)));
        regions.add(new Region("EMEA", Arrays.asList(15, 16, 17)));
        return regions;
    }

    /**
     * Gets the string value from a particular <code>Cell</code> on the given
     * <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The string value, as a <code>String</code>.
     */
    public static String getStringCellValue(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
                return c.getStringCellValue();
        }
        return null;
    }

    /**
     * Gets the boolean value from a particular <code>Cell</code> on the given
     * <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The boolean value.
     * @since 0.4.0
     */
    public static boolean getBooleanCellValue(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
                return c.getBooleanCellValue();
        }
        return false;
    }

    /**
     * Gets the <code>RichTextString</code> value from a particular
     * <code>Cell</code> on the given <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The <code>RichTextStringValue</code>.
     * @since 0.2.0
     */
    public static RichTextString getRichTextStringCellValue(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
                return c.getRichStringCellValue();
        }
        return null;
    }

    /**
     * Gets the numeric value from a particular <code>Cell</code> on the given
     * <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The numeric value, as a <code>double</code>.
     */
    public static double getNumericCellValue(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
                return c.getNumericCellValue();
        }
        return Double.NaN;
    }

    /**
     * Gets the string formula value from a particular <code>Cell</code> on the
     * given <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The string formula value.
     */
    public static String getFormulaCellValue(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                // The HSSF Formula Parser strips whitespaces outside of strings, but the
                // XSSF Formula Parser does not.
                // Just lose the spaces in all strings for comparison purposes.
                return c.getCellFormula().replaceAll("\\s+", "");
            }
        }
        return null;
    }

    /**
     * Determines whether the <code>Cell</code> on the given <code>Sheet</code>
     * at the given row and column indexes is blank: either it doesn't exist, or
     * it exists and the cell type is blank.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return Whether the <code>Cell</code> is blank.
     */
    public static boolean isCellBlank(Sheet sheet, int row, int col)
    {
        return SheetUtil.isCellBlank(sheet, row, col);
    }

    /**
     * Gets the <code>Comment</code> value, if any, from a particular
     * <code>Cell</code> on the given <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The string formula value.
     * @since 0.2.0
     */
    public static Comment getComment(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
                return c.getCellComment();
        }
        return null;
    }

    /**
     * Determines whether the <code>CellRangeAddress</code>, representing a
     * "merged region", exists in the given <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param region A <code>CellRangeAddress</code>.
     * @return <code>true</code> if the given region exists in the given sheet,
     *    <code>false</code> otherwise.
     */
    public static boolean isMergedRegionPresent(Sheet sheet, CellRangeAddress region)
    {
        int numMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numMergedRegions; i++)
        {
            CellRangeAddress candidate = sheet.getMergedRegion(i);
            if (candidate.getFirstRow() == region.getFirstRow() &&
                    candidate.getLastRow() == region.getLastRow() &&
                    candidate.getFirstColumn() == region.getFirstColumn() &&
                    candidate.getLastColumn() == region.getLastColumn())
            {
                return true;
            }
        }
        return false;
    }

    /**
     * Helper method to get an actual <code>Font</code>, regardless of which
     * kind of <code>Workbook</code> it came from.
     * @param result The result of a call to
     *    <code>RichTextStringUtil.getFontAtIndex</code>.
     * @param workbook A <code>Workbook</code>.
     * @return A <code>Font</code>.
     */
    public static Font convertToFont(Object result, Workbook workbook)
    {
        Font font;
        if (workbook instanceof HSSFWorkbook)
        {
            font = workbook.getFontAt((Short) result);
        }
        else
        {
            font = (XSSFFont) result;
        }
        return font;
    }

    /**
     * Returns the cell foreground color, as a hex string, on the given
     * <code>Sheet</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The cell foreground color, as a hex string.
     * @since 0.2.0
     */
    public static String getCellForegroundColorString(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                Color color = c.getCellStyle().getFillForegroundColorColor();
                return SheetUtil.getColorHexString(color);
            }
        }
        return null;
    }

    /**
     * Returns the cell background color, as a hex string, on the given
     * <code>Sheet</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The cell background color, as a hex string.
     * @since 0.4.0
     */
    public static String getCellBackgroundColorString(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                Color color = c.getCellStyle().getFillBackgroundColorColor();
                return SheetUtil.getColorHexString(color);
            }
        }
        return null;
    }

    /**
     * Returns the cell fill pattern on the given <code>Sheet</code>, at the
     * given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The cell fill pattern.
     * @since 0.3.0
     */
    public static FillPatternType getCellFillPattern(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                return c.getCellStyle().getFillPattern();
            }
        }
        return FillPatternType.NO_FILL;
    }

    /**
     * Returns the font color, as a hex string, on the given
     * <code>Font</code>.
     * @param workbook The <code>Workbook</code> on which the <code>Font</code>
     *    is found.
     * @param font A <code>Font</code>.
     * @return The font color, as a hex string.
     */
    public static String getFontColorString(Workbook workbook, Font font)
    {
        Color color;
        if (font instanceof HSSFFont)
        {
            color = ((HSSFFont) font).getHSSFColor((HSSFWorkbook) workbook);
        }
        else if (font instanceof XSSFFont)
        {
            color = ((XSSFFont) font).getXSSFColor();
        }
        else
        {
            throw new IllegalArgumentException("Unexpected type of Font: " + font.getClass().getName());
        }
        return SheetUtil.getColorHexString(color);
    }

    /**
     * Returns the cell bottom border color, as a hex string, on the given
     * <code>Sheet</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The cell bottom border color, as a hex string.
     * @since 0.4.0
     */
    public static String getCellBottomBorderColorString(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                CellStyle cs = c.getCellStyle();
                Color color;
                if (cs instanceof HSSFCellStyle)
                {
                    color = ExcelColor.getHssfColorByIndex(cs.getBottomBorderColor());
                }
                else if (cs instanceof XSSFCellStyle)
                {
                    color = ((XSSFCellStyle) cs).getBottomBorderXSSFColor();
                }
                else
                {
                    throw new IllegalArgumentException("Unexpected type of CellStyle for cell on sheet " +
                            sheet.getSheetName() + ", row " + row + ", col " + col);
                }
                return SheetUtil.getColorHexString(color);
            }
        }
        return null;
    }

    /**
     * Returns the cell left border color, as a hex string, on the given
     * <code>Sheet</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The cell left border color, as a hex string.
     * @since 0.4.0
     */
    public static String getCellLeftBorderColorString(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                CellStyle cs = c.getCellStyle();
                Color color;
                if (cs instanceof HSSFCellStyle)
                {
                    color = ExcelColor.getHssfColorByIndex(cs.getLeftBorderColor());
                }
                else if (cs instanceof XSSFCellStyle)
                {
                    color = ((XSSFCellStyle) cs).getLeftBorderXSSFColor();
                }
                else
                {
                    throw new IllegalArgumentException("Unexpected type of CellStyle for cell on sheet " +
                            sheet.getSheetName() + ", row " + row + ", col " + col);
                }
                return SheetUtil.getColorHexString(color);
            }
        }
        return null;
    }

    /**
     * Returns the cell right border color, as a hex string, on the given
     * <code>Sheet</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The cell right border color, as a hex string.
     * @since 0.4.0
     */
    public static String getCellRightBorderColorString(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                CellStyle cs = c.getCellStyle();
                Color color;
                if (cs instanceof HSSFCellStyle)
                {
                    color = ExcelColor.getHssfColorByIndex(cs.getRightBorderColor());
                }
                else if (cs instanceof XSSFCellStyle)
                {
                    color = ((XSSFCellStyle) cs).getRightBorderXSSFColor();
                }
                else
                {
                    throw new IllegalArgumentException("Unexpected type of CellStyle for cell on sheet " +
                            sheet.getSheetName() + ", row " + row + ", col " + col);
                }
                return SheetUtil.getColorHexString(color);
            }
        }
        return null;
    }

    /**
     * Returns the cell top border color, as a hex string, on the given
     * <code>Sheet</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The cell top border color, as a hex string.
     * @since 0.4.0
     */
    public static String getCellTopBorderColorString(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                CellStyle cs = c.getCellStyle();
                Color color;
                if (cs instanceof HSSFCellStyle)
                {
                    color = ExcelColor.getHssfColorByIndex(cs.getTopBorderColor());
                }
                else if (cs instanceof XSSFCellStyle)
                {
                    color = ((XSSFCellStyle) cs).getTopBorderXSSFColor();
                }
                else
                {
                    throw new IllegalArgumentException("Unexpected type of CellStyle for cell on sheet " +
                            sheet.getSheetName() + ", row " + row + ", col " + col);
                }
                return SheetUtil.getColorHexString(color);
            }
        }
        return null;
    }

    /**
     * Returns the <code>Hyperlink</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The <code>Hyperlink</code>, or <code>null</code> if it doesn't
     *    exist.
     * @since 0.5.0
     */
    public static Hyperlink getHyperlink(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
            {
                return c.getHyperlink();
            }
        }
        return null;
    }

    /**
     * Returns the <code>Cell</code> (if any), on the given <code>Sheet</code>,
     * at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The <code>Cell</code> or <code>null</code> if it doesn't exist.
     */
    public static Cell getCell(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
            return r.getCell(col);
        return null;
    }

    /**
     * Returns the <code>CellStyle</code> (if any), on the given
     * <code>Sheet</code>, at the given row and column indexes.
     * @param sheet The <code>Sheet</code>.
     * @param row The 0-based row index.
     * @param col The 0-based column index.
     * @return The <code>CellStyle</code> or <code>null</code> if the cell
     *    doesn't exist.
     * @since 0.4.0
     */
    public static CellStyle getCellStyle(Sheet sheet, int row, int col)
    {
        Row r = sheet.getRow(row);
        if (r != null)
        {
            Cell c = r.getCell(col);
            if (c != null)
                return c.getCellStyle();
        }
        return null;
    }
}
