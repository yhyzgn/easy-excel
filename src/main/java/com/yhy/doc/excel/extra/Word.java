package com.yhy.doc.excel.extra;

import lombok.*;

import java.io.Serializable;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 15:14
 * version: 1.0.0
 * desc   :
 */
@Data
@ToString
@EqualsAndHashCode(of = "name")
@NoArgsConstructor
@RequiredArgsConstructor
public class Word implements Serializable, Comparable {
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
    public int compareTo(Object o) {
        if (this == o) {
            return 0;
        }
        if (this.name == null) {
            return -1;
        }
        if (o == null) {
            return 1;
        }
        if (!(o instanceof Word)) {
            return 1;
        }
        String t = ((Word) o).getName();
        if (t == null) {
            return 1;
        }
        return this.name.compareTo(t);
    }
}
