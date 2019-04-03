package net.sf.jett.model;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * A <code>HashMapWrapper</code> is a <code>HashMap</code> that "wraps" another
 * <code>HashMap</code>.  All mappings in this map "override" any mappings in
 * the wrapped map.  Many <code>HashMapWrappers</code> can be used to wrap a
 * single map, to avoid the need to clone a single map multiple times.  This
 * wrapper doesn't necessarily override all <code>Map</code> methods.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class HashMapWrapper<K, V> extends HashMap<K, V>
{
    private Map<K, V> myWrappedMap;
    private int mySize = 0;

    /**
     * Constructs a <code>HashMapWrapper</code> that wraps the given <code>Map</code>.
     * @param otherMap Another <code>Map</code> to wrap.
     */
    public HashMapWrapper(Map<K, V> otherMap)
    {
        myWrappedMap = otherMap;
        mySize = myWrappedMap.size();
    }

    /**
     * Clears all entries in this map and in the wrapped map.
     */
    @Override
    public void clear()
    {
        super.clear();
        myWrappedMap.clear();
        mySize = 0;
    }

    /**
     * Looks in this map, and if not found, looks in the wrapped map.
     * @param key A key that may be present in this map or in the wrapped map.
     * @return If this map or the wrapped map contains the key.
     */
    @Override
    public boolean containsKey(Object key)
    {
        return super.containsKey(key) || myWrappedMap.containsKey(key);
    }

    /**
     * Looks in this map, and if not found, looks in the wrapped map.
     * @param value A value that may be present in this map or in the wrapped map.
     * @return If this map or the wrapped map contains the value.
     */
    @Override
    public boolean containsValue(Object value)
    {
        return super.containsValue(value) || myWrappedMap.containsValue(value);
    }

    /**
     * Returns a <code>Set</code> of all entries in this map and the wrapped
     * map, with this map overriding any entries in the wrapped map.  Changes
     * to the returned <code>Set</code> do not write through to this map.
     * @return A <code>Set</code> of mappings in this map and the wrapped map.
     */
    @Override
    public Set<Map.Entry<K, V>> entrySet()
    {
        Set<Map.Entry<K, V>> entrySet = new HashSet<>(super.entrySet());
        entrySet.addAll(myWrappedMap.entrySet());
        return entrySet;
    }

    /**
     * Returns the value from this map to which the key is mapped, or from the
     * wrapped map if not present in this map.
     * @param key The key.
     * @return The value, if the key is present in either this map or the
     *    wrapped map, or <code>null</code> if not present.
     */
    @Override
    public V get(Object key)
    {
        if (super.containsKey(key))
        {
            return super.get(key);
        }
        return myWrappedMap.get(key);
    }

    /**
     * Returns <code>true</code> if this map and the wrapped map are both empty.
     * @return <code>true</code> if this map and the wrapped map are both empty.
     */
    @Override
    public boolean isEmpty()
    {
        return super.isEmpty() && myWrappedMap.isEmpty();
    }

    /**
     * Returns a <code>Set</code> of all keys in this map and the wrapped
     * map, with this map overriding any keys in the wrapped map.  Changes
     * to the returned <code>Set</code> do not write through to this map.
     * @return A <code>Set</code> of keys in this map and the wrapped map.
     */
    @Override
    public Set<K> keySet()
    {
        Set<K> keySet = new HashSet<>(super.keySet());
        keySet.addAll(myWrappedMap.keySet());
        return keySet;
    }

    /**
     * Maps the given key to the given value in this map, never the wrapped map.
     * @param key The key to map.
     * @param value The value to map.
     * @return The old value, whether it came from this map or the wrapped map.
     */
    @Override
    public V put(K key, V value)
    {
        if (!super.containsKey(key) && !myWrappedMap.containsKey(key))
            mySize++;
        V oldValue;
        if (super.containsKey(key))
        {
            oldValue = super.put(key, value);
        }
        else
        {
            oldValue = myWrappedMap.get(key);
            super.put(key, value);
        }
        return oldValue;
    }

    /**
     * Puts all entries from the given map into this map, never the wrapped map.
     * @param map Another map.
     */
    @Override
    public void putAll(Map<? extends K, ? extends V> map)
    {
        for (Map.Entry<? extends K, ? extends V> entry : map.entrySet())
        {
            put(entry.getKey(), entry.getValue());
        }
    }

    /**
     * Removes the entry associated with this key from this map, never the
     * wrapped map.
     * @param key The key associated with the entry to remove.
     * @return The value that was associated from this map, never the wrapped
     *    map.
     */
    @Override
    public V remove(Object key)
    {
        if (super.containsKey(key))
        {
            if (!myWrappedMap.containsKey(key))
            {
                mySize--;
            }
            return super.remove(key);
        }
        return null;
    }

    /**
     * Returns the number of mappings in this map unioned with the wrapped map.
     * Any keys present in both maps are counted only once.
     * @return The number of mappings.
     */
    @Override
    public int size()
    {
        return mySize;
    }

    /**
     * Returns a <code>Collection</code> of all values in this map and the
     * wrapped map, with this map overriding any values in the wrapped map with
     * the same key.  Changes to the returned <code>Collection</code> do not
     * write through to this map.
     * @return A <code>Collection</code> of values in this map and the wrapped
     *    map.
     */
    @Override
    public Collection<V> values()
    {
        List<V> values = new ArrayList<>();
        for (Map.Entry<K, V> entry : entrySet())
        {
            values.add(entry.getValue());
        }
        return values;
    }
}
