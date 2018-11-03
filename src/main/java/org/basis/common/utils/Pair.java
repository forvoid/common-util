package org.basis.common.utils;

public class Pair<K, V> {
    K k;
    V v;

    private Pair(K k, V v) {
        this.k = k;
        this.v = v;
    }

    public static <K, V> Pair<K, V> create(K k, V v) {
        return new Pair(k, v);
    }

    public K getFirst() {
        return k;
    }

    public V getSecond() {
        return v;
    }
}