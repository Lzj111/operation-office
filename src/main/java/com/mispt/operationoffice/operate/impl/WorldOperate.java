package com.mispt.operationoffice.operate.impl;

import com.mispt.operationoffice.entity.ReplaceValue;
import com.mispt.operationoffice.entity.SeriesData;
import com.mispt.operationoffice.utils.ChartUtil;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

/**
 * @Classname WorldOperate
 * @Description
 * @Date 2022/12/20 15:39
 * @Author by lzj
 */
@Component("worldOperate")
public class WorldOperate extends BaseOperate {

    @Override
    public void execLogic(String filePath) throws IOException {
        // 1> 构建word对象
        FileInputStream fs = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(fs);

        // 2> 获取所有段落
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        paragraphs.forEach(this::handleParagraph);

        // 3> 获取所有表格
        List<XWPFTable> tables = doc.getTables();
        tables.forEach(this::handleTable);

        // 4> 获取所有图表
        List<XWPFChart> charts = doc.getCharts();
        charts.forEach(this::handleChart);

        // n> 持久化修改至PPT
        FileOutputStream fo = new FileOutputStream(filePath);
        doc.write(fo);
        fo.flush();
        fo.close();
    }

    /**
     * 操作段落对象
     * @param paragraph
     */
    public void handleParagraph(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (null == runs || runs.size() == 0) {
            return;
        }

        // 1> 循环段落
        runs.forEach(item -> {
            // 获取段落里的文本内容
            String text = item.getText(0);
            System.out.println("段落文本：" + text);

            if (!StringUtils.isEmpty(text)) {
                List<ReplaceValue> textList = dataReplace.getText();
                for (int i = 0; i < textList.size(); i++) {
                    ReplaceValue data = textList.get(i);
                    String code = data.getCode();
                    String value = data.getValue().toString();
                    // >替换文本内容
                    if (text.contains(code)) {
                        text = text.replace(code, value);
                        // 把替换好的文本内容，保存到当前这个文本对象
                        item.setText(text, 0);
                    }
                }
            }
        });
    }

    /**
     * 操作表格对象
     * @param table
     */
    public void handleTable(XWPFTable table) {
        List<ReplaceValue> datas = dataReplace.getTable();
        List<XWPFTableRow> tableRows = table.getRows();
        // 1 > 循环所有行
        StringBuilder sb = new StringBuilder(">> 表格组件的内容：\n");
        for (int i = 0; i < tableRows.size(); i++) {
            XWPFTableRow row = tableRows.get(i);

            // 2> 循环当前行的所有列
            List<XWPFTableCell> cells = row.getTableCells();
            for (int j = 0; j < cells.size(); j++) {
                // >> 获取单元格文本值
                XWPFTableCell cell = cells.get(j);
                String cellTxt = cell.getText();
                if (StringUtils.isEmpty(cellTxt)) {
                    continue;
                }
                // >> 从配置中读取替换对象
                ReplaceValue replaceValue = ChartUtil.getReplaceValueContainsKey(datas, cellTxt);
                if (null == replaceValue) {
                    continue;
                }

                // 3> 根据是否匹配到配置来替换值
                // 注意，getParagraphs一定不能漏掉,因为一个表格里面可能会有多个需要替换的文字,如果没有这个步骤那么文字会替换不了
                String replTxt = replaceValue.getValue().toString();
                for (XWPFParagraph p : cell.getParagraphs()) {
                    for (XWPFRun r : p.getRuns()) {
                        r.setText(replTxt, 0);
                    }
                }
                cell.setColor("cfabff");
                sb.append("\t" + cellTxt + "->" + replTxt);
            }
            sb.append("\n");
        }
        System.out.println(sb);
    }

    /**
     * 操作图表对象
     * @param chart
     */
    public void handleChart(XWPFChart chart) {
        try {
            // 1> 准备数据(系列数据)
            List<SeriesData> seriesDatas = ChartUtil.getSeriesDataList();

            // 2> 调用更新图表数据(excel)
            CTPlotArea plot = chart.getCTChart().getPlotArea();
            XSSFWorkbook workbook = chart.getWorkbook();
//            ChartUtil.updateChartData(seriesDatas, plot, workbook, chart);
            List<ReplaceValue> datas = dataReplace.getChart();
            ChartUtil.updateChartDataConfig(datas, plot, workbook, chart);
        } catch (Exception e) {
            logger.error("处理图表类型的PPT组件异常：", e);
        }
    }
}
