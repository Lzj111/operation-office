package com.mispt.operationoffice.operate.impl;

import com.mispt.operationoffice.entity.DataReplace;
import com.mispt.operationoffice.entity.ReplaceValue;
import com.mispt.operationoffice.entity.SeriesData;
import com.mispt.operationoffice.utils.ChartUtil;
import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

/**
 * @Description 操作PPT对象
 * @Author lzj
 * @Date 2022/12/22
 **/
@Component("pptOperate")
public class PPTOperate extends BaseOperate {

    protected final Logger logger = LoggerFactory.getLogger(this.getClass());

    @Override
    public void execLogic(String filePath) throws IOException {
        // 1> 构建ppt对象
        FileInputStream fs = new FileInputStream(filePath);
        XMLSlideShow xmlSlideShow = new XMLSlideShow(fs);
        fs.close();

        // 2> 获取所有的幻灯片(页码)
        List<XSLFSlide> slides = getAllSlides(xmlSlideShow);
        slides.forEach(this::handleShape);

        // 3> 操作图表数据
        List<XSLFChart> charts = getAllCharts(xmlSlideShow);
        charts.forEach(chart -> {
            System.out.println("> 当前组件类型:" + chart.getClass());
            handleChartType(chart);
        });

        // n> 持久化修改至PPT
        FileOutputStream fo = new FileOutputStream(filePath);
        xmlSlideShow.write(fo);
        fo.flush();
        fo.close();
    }

    /**
     * 获取所有的幻灯片(页码)
     * @param xmlSlideShow
     * @return
     */
    public List<XSLFSlide> getAllSlides(XMLSlideShow xmlSlideShow) {
        // 获取ppt中所有幻灯片；4.x版本以下poi-ooxml是数组
        return xmlSlideShow.getSlides();
    }

    /**
     * 获取所有的图表
     * @param xmlSlideShow
     * @return
     */
    public List<XSLFChart> getAllCharts(XMLSlideShow xmlSlideShow) {
        // 获取所有的图表
        return xmlSlideShow.getCharts();
    }

    /**
     * 循环当前slide下的组件
     * @param slide
     */
    public void handleShape(XSLFSlide slide) {
        // 获取每张幻灯片中的shape(形状)
        List<XSLFShape> shapes = slide.getShapes();
        shapes.forEach(shape -> {
            System.out.println("> 当前组件类型:" + shape.getClass());
            // 文本类型组件
            if (shape instanceof XSLFTextShape) {
                handleTextType(shape);
            }
            // 表格类型组件
            else if (shape instanceof XSLFTable) {
                handleTableType(shape);
            }
            // 图表类型的组件(图表类型的单独处理)
            else if (shape instanceof XSLFGraphicFrame) {
            }
            // 图片类型的组件
            else if (shape instanceof XSLFPictureShape) {
                handlePictureType(shape);
            }
        });
    }

    /**
     * 处理文本类型的PPT组件
     * @param shape
     */
    public void handleTextType(XSLFShape shape) {
        XSLFTextShape txShape = (XSLFTextShape) shape;
        String txtContent = txShape.getText();
        if (StringUtils.isEmpty(txtContent)) {
            return;
        }

        List<ReplaceValue> textList = dataReplace.getText();
        textList.forEach(data -> {
            String code = data.getCode();
            String value = data.getValue().toString();
            // 文本包含替换code
            if (txtContent.contains(code)) {
                System.out.println(">> 文本组件的内容：" + txtContent + ";" + code + "->" + value);
                txShape.setText(txtContent.replace(code, value));
            }
        });
    }

    /**
     * 处理表格类型的PPT组件
     * @param shape
     */
    public void handleTableType(XSLFShape shape) {
        StringBuilder sb = new StringBuilder();
        XSLFTable table = (XSLFTable) shape;
        List<XSLFTableRow> rows = table.getRows();
        List<ReplaceValue> datas = dataReplace.getTable();
        // > 循环行
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            XSLFTableRow row = rows.get(rowIndex);
            List<XSLFTableCell> cells = row.getCells();
            // >> 循环列
            for (int cellIndex = 0; cellIndex < cells.size(); cellIndex++) {
                XSLFTableCell cell = cells.get(cellIndex);
                String cellTxt = cell.getText();
                // >>> 从配置中读取替换对象
                ReplaceValue replaceValue = ChartUtil.getReplaceValueContainsKey(datas, cellTxt);
                if (null == replaceValue) {
                    continue;
                }

                // >>> 替换文本,设置颜色
                sb.append("\t").append(cellTxt).append((cellIndex > 0 && rowIndex > 0) ? "->" + replaceValue.getValue() : "");
                cell.setText(String.valueOf(replaceValue.getValue()));
                cell.setFillColor(new Color(207, 171, 255));
            }
            sb.append("\n");
        }

        System.out.println(">> 表格组件的内容：\n" + sb);
    }

    /**
     * 处理图表类型的PPT组件
     * @param chart
     */
    public void handleChartType(XSLFChart chart) {
        try {
            // 1> 获取图表标题
            String text = chart.getTitleShape().getText();
            System.out.println(">> 图表名称：" + chart.getTitleShape().getText() + ",标题：" + text);

            // 2> 准备数据(系列数据)
            List<SeriesData> seriesDatas = ChartUtil.getSeriesDataList();

            // 3> 调用更新图表数据(excel)
            CTPlotArea plot = chart.getCTChart().getPlotArea();
            XSSFWorkbook workbook = chart.getWorkbook();
            // 3.1> 替换所有数据
            ChartUtil.updateChartData(seriesDatas, plot, workbook, chart);
            // 3.2> 替换配置数据
//            List<ReplaceValue> datas = dataReplace.getChart();
//            ChartUtil.updateChartDataConfig(datas, plot, workbook, chart);
        } catch (Exception e) {
            logger.error("处理图表类型的PPT组件异常：", e);
        }
    }

    /**
     * 处理图片类型的PPT组件
     * @param shape
     */
    public void handlePictureType(XSLFShape shape) {
//        // 获取所有图表中系列；此处概念需要了解PPT的图表结构
//        // 获取图表的图表区域
//        CTPlotArea plotArea = chart.getCTChart().getPlotArea();
//        // 获取区域中的柱状图
//        CTBarChart barchart = plotArea.getBarChartArray(0);
//        // 获取柱状图的序列
//        List<CTBarSer> serList = barchart.getSerList();
//
//        CTBarSer ctBarSer = serList.get(0);
//        // 获取系列名称
//        String serName = ctBarSer.getTx().getStrRef().getF();
//        // 系列对象CTBarSer 可以操作图表数据
    }

}
