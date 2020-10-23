package com.yhy.doc.excel.extra;

import com.hankcs.hanlp.HanLP;
import com.hankcs.hanlp.seg.common.Term;
import lombok.extern.slf4j.Slf4j;

import java.util.List;
import java.util.stream.Collectors;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 15:10
 * version: 1.0.0
 * desc   :
 */
@Slf4j
public class Segmenter {

    /**
     * 中文句子分词
     *
     * @param sentence 句子
     * @return 分词结果
     */
    public static List<Word> segment(String sentence) {
        List<Term> terms = HanLP.segment(sentence);
        if (null != terms) {
            return terms.stream().map(term -> new Word(term.word, term.nature.toString())).collect(Collectors.toList());
        }
        return null;
    }
}
