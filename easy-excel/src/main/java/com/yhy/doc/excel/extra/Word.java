package com.yhy.doc.excel.extra;

import lombok.*;
import org.jetbrains.annotations.NotNull;

import java.io.Serializable;

/**
 * 分词信息
 * <p>
 * Created on 2019-09-09 15:14
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Data
@ToString
@EqualsAndHashCode(of = "name")
@NoArgsConstructor
@RequiredArgsConstructor
public class Word implements Serializable, Comparable<Word> {
    private static final long serialVersionUID = -5647329507891444116L;

    /**
     * 词名
     */
    @NonNull
    private String name;

    /**
     * 词性
     */
    @NonNull
    private String nature;

    /**
     * 权重
     */
    private Float weight;

    @Override
    public int compareTo(@NotNull Word o) {
        if (this == o) {
            return 0;
        }
        String t = o.getName();
        return this.name.compareTo(t);
    }
}
