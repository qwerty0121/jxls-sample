package com.qwerty0121.jxls.sample;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.jxls.builder.JxlsOutputFile;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

import com.qwerty0121.jxls.sample.bean.Detail;
import com.qwerty0121.jxls.sample.command.PrintAreaCommand;

public class JxlsSample {

    public static void main(String[] args) {
        try {
            outputReport();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 帳票を出力する
     * 
     * @throws IOException
     */
    private static void outputReport() throws IOException {
        try (var template = loadTemplate();) {
            // 可変データを作成
            var data = generateVariableData();

            // 出力先
            var output = new File("./.output/report.xlsx");

            // 帳票出力
            JxlsPoiTemplateFillerBuilder.newInstance()
                    .withTemplate(template)
                    .withCommand(PrintAreaCommand.COMMAND_NAME, PrintAreaCommand.class)
                    .build()
                    .fill(data, new JxlsOutputFile(output));
        }
    }

    /**
     * 可変データを生成する
     * 
     * @return 可変データ
     */
    private static Map<String, Object> generateVariableData() {
        var data = new HashMap<String, Object>();

        data.put("text", "test text");
        data.put("number", 12345);
        data.put("date", LocalDate.now());

        var details = IntStream.rangeClosed(1, 10).mapToObj(index -> {
            var detail = new Detail();
            detail.setId(index);
            detail.setName(String.format("テストユーザー%d", index));
            detail.setAmount(new BigDecimal(index * 1000));
            return detail;
        }).collect(Collectors.toList());
        data.put("details", details);

        return data;
    }

    /**
     * テンプレートファイルを読み込む
     * 
     * @return テンプレートファイル
     */
    private static InputStream loadTemplate() {
        return JxlsSample.class.getClassLoader().getResourceAsStream("templates/report/template.xlsx");
    }

}
