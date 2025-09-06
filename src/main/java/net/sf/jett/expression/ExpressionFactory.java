package net.sf.jett.expression;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlExpression;
import org.apache.commons.jexl3.internal.introspection.Permissions;
import org.apache.commons.jexl3.introspection.JexlPermissions;

/**
 * <p>An <code>ExpressionFactory</code> is a factory class that
 * creates and uses a <code>JexlEngine</code> to create JEXL
 * <code>Expressions</code>.</p>
 *
 * <p>It passes through several items to the JEXL Engine, including "silent"
 * and "lenient" flags, the JEXL parse cache size, and JEXL namespace function
 * objects (including jAgg functionality).</p>
 *
 * <p>As of 0.9.0, this class is no longer a singleton, to allow concurrent
 * <code>ExcelTransformers</code> to avoid contention by having their own
 * <code>ExpressionFactories</code>.
 *
 * @author Randy Gettman
 */
public class ExpressionFactory
{
    private JexlEngine myEngine;
    private Map<String, Object> myFuncs = new HashMap();
    private boolean isStrict = false;
    private boolean isDebug = false;
    private boolean isSilent = true;
    private int cacheSize = 512;
    private JexlPermissions permissions = JexlPermissions.parse(null); // DEFAULT | null (srcs) > UNRESTRICTED!
    
    public ExpressionFactory() {
        myFuncs.put("jagg", JaggFuncs.class);
        myFuncs.put("jett", JettFuncs.class);
    }
    
    /**
     * Constructs a <code>ExpressionFactory</code>.  Initializes an internal
     * <code>JexlEngine</code> and initializes the functions map.
     */
    public boolean isDebug() {
        return isDebug;
    }
    public boolean isStrict() {
        return isStrict;
    }
    public boolean isSilent() {
        return isSilent;
    }
    public void setDebug(boolean val) { if (isDebug == val) return;isDebug=val;myEngine=null; }
    public void setStrict(boolean val) { if (isStrict == val) return;isStrict=val;myEngine=null; }
    public void setSilent(boolean val) { if (isSilent == val) return;isSilent=val;myEngine=null; }
    public void setCacheSize(int val) { if (cacheSize == val) return;cacheSize=val;myEngine=null; }

    private void check() {
        if (myEngine != null)
            return;
        myFuncs.put("jagg", JaggFuncs.class);
        myFuncs.put("jett", JettFuncs.class);
        myEngine=new JexlBuilder().strict(isStrict).debug(isDebug).silent(isSilent).cache(cacheSize).namespaces(myFuncs).permissions(permissions).create();
    }

    public JexlExpression createExpression(String expression)  {
        check();
        return myEngine.createExpression(expression);
    }
    public void registerFuncs(String namespace, Object funcsObject) {
        if (myFuncs.containsKey(namespace))
            throw new IllegalArgumentException("ExpressionFactory : namespace "+namespace+" already registed!");
        myFuncs.put(namespace, funcsObject);
        myEngine=null;
    }
    
    /* DEFAULT IS ALLOW ALL 
     ["java.util.*","!java.lang.reflect.**","java.exact.class"]..
     */
    public void permissions(String rules[]) {
        permissions = JexlPermissions.parse(rules);
        myEngine=null;
    }
    
    public JexlEngine createJexlEngine() {
        return new JexlBuilder().strict(isStrict).debug(isDebug).silent(isSilent).cache(cacheSize).namespaces(myFuncs).permissions(permissions).create();
    }
}

