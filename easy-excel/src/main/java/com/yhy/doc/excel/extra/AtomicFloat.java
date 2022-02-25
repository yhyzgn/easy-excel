package com.yhy.doc.excel.extra;

import java.util.concurrent.atomic.AtomicInteger;

/**
 * float 类型的原子操作
 * <p>
 * Created on 2019-09-09 15:29
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public class AtomicFloat extends Number {
    private static final long serialVersionUID = -5823759557708837608L;

    private final AtomicInteger bits;

    public AtomicFloat() {
        this(0f);
    }

    public AtomicFloat(float bits) {
        this.bits = new AtomicInteger(Float.floatToIntBits(bits));
    }

    public final float addAndGet(float delta) {
        float expect, update;
        do {
            expect = get();
            update = expect + delta;
        } while (!this.compareAndSet(expect, update));
        return update;
    }

    public final float getAndAdd(float delta) {
        float expect, update;
        do {
            expect = get();
            update = expect + delta;
        } while (!this.compareAndSet(expect, update));
        return expect;
    }

    public final float incrementAndGet() {
        return addAndGet(1);
    }

    public final float getAndIncrement() {
        return getAndAdd(1);
    }

    public final float decrementAndGet() {
        return addAndGet(-1);
    }

    public final float getAndDecrement() {
        return getAndAdd(-1);
    }

    public final boolean compareAndSet(float expect, float update) {
        return bits.compareAndSet(Float.floatToIntBits(expect), Float.floatToIntBits(update));
    }

    public final void set(float value) {
        bits.set(Float.floatToIntBits(value));
    }

    public final float get() {
        return Float.intBitsToFloat(bits.get());
    }

    @Override
    public int intValue() {
        return (int) get();
    }

    @Override
    public long longValue() {
        return (long) get();
    }

    @Override
    public float floatValue() {
        return get();
    }

    @Override
    public double doubleValue() {
        return floatValue();
    }
}
