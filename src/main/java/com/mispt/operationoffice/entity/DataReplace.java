package com.mispt.operationoffice.entity;

import java.util.List;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * @Classname DataReplace
 * @Description
 * @Date 2022/12/26 10:02
 * @Author by lzj
 */
@Component
@ConfigurationProperties(prefix = "data-replace")
public class DataReplace {

    private List<ReplaceValue> text;
    private List<ReplaceValue> chart;
    private List<ReplaceValue> table;

    public List<ReplaceValue> getText() {
        return text;
    }

    public void setText(List<ReplaceValue> text) {
        this.text = text;
    }

    public List<ReplaceValue> getChart() {
        return chart;
    }

    public void setChart(List<ReplaceValue> chart) {
        this.chart = chart;
    }

    public List<ReplaceValue> getTable() {
        return table;
    }

    public void setTable(List<ReplaceValue> table) {
        this.table = table;
    }
}

