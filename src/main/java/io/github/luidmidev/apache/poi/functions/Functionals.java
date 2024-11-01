package io.github.luidmidev.apache.poi.functions;

import java.util.function.BiConsumer;
import java.util.function.Consumer;

public final class Functionals {

    private Functionals() {
        throw new UnsupportedOperationException("Utility class");
    }

    public static  <T, U> BiConsumer<T, U> biConsumerNoAction() {
        return (t, u) -> {
        };
    }

    public static  <T> Consumer<T> consumerNoAction() {
        return t -> {
        };
    }

}
