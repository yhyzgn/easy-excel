package com.yhy.doc.excel.internal;

import com.yhy.doc.excel.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 15:41
 * version: 1.0.0
 * desc   : 根据余弦夹角定理求字符串相似度
 * <p>
 * 判定方式：余弦相似度，通过计算两个向量的夹角余弦值来评估他们的相似度 余弦夹角原理：
 * - a•b = |a||b|cosθ
 * - 向量 a=(x1,y1), 向量 b=(x2,y2) 则： similarity = a.b/|a|*|b|  a.b=x1x2+y1y2
 * - 其中 |a| = 根号[(x1)^2+(y1)^2], |b| = 根号[(x2)^2+(y2)^2]
 */
@Slf4j
public class CosineSimilarity {

    /**
     * 用余弦夹角公式计算相似度
     *
     * @param text1 字符串1
     * @param text2 字符串2
     * @return 相似度
     */
    public static double getSimilarity(String text1, String text2) {
        // 如果两个字符串都为空，就完全相同
        if (StringUtils.isEmpty(text1) && StringUtils.isEmpty(text2)) {
            return 1.0;
        }
        // 如果只有一个为空，就完全不同
        if (StringUtils.isEmpty(text1) || StringUtils.isEmpty(text2)) {
            return 0;
        }
        // 如果两个字符串完全相等，就完全相同
        if (text1.equalsIgnoreCase(text2)) {
            return 1.0;
        }

        List<Word> words1 = Segmenter.segment(text1);
        List<Word> words2 = Segmenter.segment(text2);

        return getSimilarity(words1, words2);
    }

    /**
     * 用余弦夹角公式计算相似度
     *
     * @param words1 分词集合1
     * @param words2 分词结合2
     * @return 相似度
     */
    public static double getSimilarity(List<Word> words1, List<Word> words2) {
        double similarity = getCosineSimilarity(words1, words2);
        // (int) (score * 1000000 + 0.5)其实代表保留小数点后六位 ,因为1034234.213强制转换不就是1034234。对于强制转换添加0.5就等于四舍五入
        similarity = (int) (similarity * 1000000 + 0.5) / (double) 1000000;
        return similarity;
    }

    /**
     * 用余弦夹角公式计算相似度
     *
     * @param words1 分词集合1
     * @param words2 分词结合2
     * @return 相似度
     */
    private static double getCosineSimilarity(List<Word> words1, List<Word> words2) {
        // 计算词频，也就是权重
        computeWeightByFrequency(words1, words2);

        // 分别获取到上一步计算好的词频
        Map<String, Float> weightMap1 = mapFrequency(words1);
        Map<String, Float> weightMap2 = mapFrequency(words2);

        // 将所有词装入集合中
        Set<Word> wordSet = new HashSet<>();
        wordSet.addAll(words1);
        wordSet.addAll(words2);

        // 运用公式：a•b = |a||b|cosθ 来计算
        // a•b
        AtomicFloat ab = new AtomicFloat();
        // |a| * |a|
        AtomicFloat aa = new AtomicFloat();
        // |b| * |b|
        AtomicFloat bb = new AtomicFloat();

        // 计算设置词频向量
        wordSet.parallelStream().forEach(word -> {
            Float a = weightMap1.get(word.getName());
            Float b = weightMap2.get(word.getName());

            // 向量ab
            if (null != a && null != b) {
                ab.addAndGet(a * b);
            }
            if (null != a) {
                aa.addAndGet(a * a);
            }
            if (null != b) {
                bb.addAndGet(b * b);
            }
        });

        // 分别计算 a,b 向量的长度
        double valueA = Math.sqrt(aa.doubleValue());
        double valueB = Math.sqrt(bb.doubleValue());

        // 用公式 cosθ = a•b / (|a|*|b|) 计算余弦值
        // 使用BigDecimal保证精确计算浮点数
        BigDecimal decimal = BigDecimal.valueOf(valueA).multiply(BigDecimal.valueOf(valueB));
        // decimal被除数，9表示小数点后保留9位，最后一个表示用标准的四舍五入法
        return BigDecimal.valueOf(ab.get()).divide(decimal, 9, RoundingMode.HALF_UP).doubleValue();
    }

    /**
     * 计算分词词频
     *
     * @param words1 分词
     * @param words2 分词
     */
    private static void computeWeightByFrequency(List<Word> words1, List<Word> words2) {
        computeWeightByFrequency(words1);
        computeWeightByFrequency(words2);
    }

    /**
     * 计算分词词频
     *
     * @param words 分词
     */
    private static void computeWeightByFrequency(List<Word> words) {
        if (CollectionUtils.isEmpty(words) || null != words.get(0).getWeight()) {
            return;
        }
        Map<String, AtomicFloat> frequency = getFrequency(words);
        if (log.isDebugEnabled()) {
            log.info("词频统计：\n{}", getWordsFrequencyString(frequency));
        }
        words.parallelStream().forEach(word -> word.setWeight(frequency.get(word.getName()).get()));
    }

    /**
     * 计算每个分词词频
     *
     * @param words 分词集合
     * @return 词频信息
     */
    private static Map<String, AtomicFloat> getFrequency(List<Word> words) {
        Map<String, AtomicFloat> frequency = new HashMap<>();
        // 按分词统计词频
        words.forEach(word -> frequency.computeIfAbsent(word.getName(), key -> new AtomicFloat()).incrementAndGet());
        return frequency;
    }

    /**
     * 将分词信息集合转换成map
     *
     * @param words 分词集合
     * @return 分词词频信息map
     */
    private static Map<String, Float> mapFrequency(List<Word> words) {
        if (CollectionUtils.isEmpty(words)) {
            return Collections.emptyMap();
        }
        Map<String, Float> frequencyMap = new ConcurrentHashMap<>(words.size());
        words.parallelStream().forEach(word -> {
            if (null != word.getWeight()) {
                frequencyMap.put(word.getName(), word.getWeight());
            } else {
                log.error("No weight of word : {}", word.getName());
            }
        });
        return frequencyMap;
    }

    /**
     * 输出词频统计
     *
     * @param frequency 词频信息
     * @return 统计信息
     */
    private static String getWordsFrequencyString(Map<String, AtomicFloat> frequency) {
        StringBuilder str = new StringBuilder();
        if (frequency != null && !frequency.isEmpty()) {
            AtomicFloat integer = new AtomicFloat();
            frequency.entrySet().stream().sorted((a, b) -> (int) (b.getValue().get() - a.getValue().get())).forEach(
                    i -> str.append("\t").append(integer.incrementAndGet()).append("、").append(i.getKey()).append("=").append(i.getValue()).append("\n"));
        }
        str.setLength(str.length() - 1);
        return str.toString();
    }
}
