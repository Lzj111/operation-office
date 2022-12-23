package com.mispt.operationoffice.operate.impl;

import com.mispt.operationoffice.entity.NameDouble;
import com.mispt.operationoffice.entity.SeriesData;
import java.awt.Color;
import java.awt.FileDialog;
import java.awt.Frame;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;
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

        String speed = "{jindu}";
        String speedStr = Math.ceil(Math.random() * 100) + "%";
        // 包含变量：{jindu} 替换
        if (txtContent.contains(speed)) {
            System.out.println(">> 文本组件的内容：" + txtContent + ";" + speed + "->" + speedStr);
            txShape.setText(txtContent.replace(speed, speedStr));
        }
    }

    /**
     * 处理表格类型的PPT组件
     * @param shape
     */
    public void handleTableType(XSLFShape shape) {
        StringBuilder sb = new StringBuilder();
        XSLFTable table = (XSLFTable) shape;
        List<XSLFTableRow> rows = table.getRows();

        // > 循环行
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            XSLFTableRow row = rows.get(rowIndex);
            List<XSLFTableCell> cells = row.getCells();
            // > 循环列
            for (int cellIndex = 0; cellIndex < cells.size(); cellIndex++) {
                XSLFTableCell cell = cells.get(cellIndex);
                String cellTxt = cell.getText();
                String replTxt = String.valueOf(Math.ceil(Math.random() * 100));
                sb.append("\t").append(cellTxt).append((cellIndex > 0 && rowIndex > 0) ? "->" + replTxt : "");

                // 更新表格内容
                if (rowIndex > 0 && cellIndex > 0) {
                    cell.setText(replTxt);
                    cell.setFillColor(new Color(207, 171, 255));
                }
            }
            sb.append("\n");
        }

        System.out.println(">> 表格组件的内容：\n" + sb);
    }

    //region 处理图表类型的PPT组件

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
            List<SeriesData> seriesDatas = Arrays.asList(
                    new SeriesData("一年级", Arrays.asList(
                            new NameDouble("2014271班", Math.random() * 100),
                            new NameDouble("2014272班", Math.random() * 100),
                            new NameDouble("2014273班", Math.random() * 100),
                            new NameDouble("2014274班", Math.random() * 100),
                            new NameDouble("2014275班", Math.random() * 100),
                            new NameDouble("2014276班", Math.random() * 100)
                    )),
                    new SeriesData("二年级", Arrays.asList(
                            new NameDouble("2014271班", Math.random() * 100),
                            new NameDouble("2014272班", Math.random() * 100),
                            new NameDouble("2014273班", Math.random() * 100),
                            new NameDouble("2014274班", Math.random() * 100),
                            new NameDouble("2014275班", Math.random() * 100),
                            new NameDouble("2014276班", Math.random() * 100)
                    )),
                    new SeriesData("三年级", Arrays.asList(
                            new NameDouble("2014271班", Math.random() * 100),
                            new NameDouble("2014272班", Math.random() * 100),
                            new NameDouble("2014273班", Math.random() * 100),
                            new NameDouble("2014274班", Math.random() * 100),
                            new NameDouble("2014275班", Math.random() * 100),
                            new NameDouble("2014276班", Math.random() * 100)
                    )),
                    new SeriesData("四年级", Arrays.asList(
                            new NameDouble("2014271班", Math.random() * 100),
                            new NameDouble("2014272班", Math.random() * 100),
                            new NameDouble("2014273班", Math.random() * 100),
                            new NameDouble("2014274班", Math.random() * 100),
                            new NameDouble("2014275班", Math.random() * 100),
                            new NameDouble("2014276班", Math.random() * 100)
                    ))
            );

            // 3>查看里面的图表数据，才能知道是什么图表
            CTPlotArea plot = chart.getCTChart().getPlotArea();
            XSSFWorkbook workbook = chart.getWorkbook();
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 4> 柱状图
            if (!plot.getBarChartList().isEmpty()) {
                // >> 更新图表Excel的数据
                updateChartExcelV(seriesDatas, workbook, sheet);

                // >> 获取c:barChart的xml对象
                CTBarChart barChart = plot.getBarChartArray(0);
                // >> 更新chart的缓存数据
                for (int i = 0; i < barChart.getSerList().size(); i++) {
                    // >>> 获取图表的一个系列对象
                    CTBarSer ser = barChart.getSerList().get(i);
                    // >>> getTx:系列的标题缓存; getCat:维度的数据缓存; getVal:数据的缓存
                    updateChartCatAndNum(seriesDatas.get(i), ser.getTx(), ser.getCat(), ser.getVal());
                }
            }
            // 饼图
            else if (!plot.getPieChartList().isEmpty()) {

            }
        } catch (Exception e) {
            logger.error("处理图表类型的PPT组件异常：", e);
        }
    }

    /**
     * 更新图表的关联 excel，值是纵向的
     * @param seriesDatas
     * @param workbook
     * @param sheet
     */
    private void updateChartExcelV(List<SeriesData> seriesDatas, XSSFWorkbook workbook, XSSFSheet sheet) {
        XSSFRow title = sheet.getRow(0);
        // > 循环替换的行系列数据
        for (int i = 0; i < seriesDatas.size(); i++) {
            SeriesData data = seriesDatas.get(i);

            // >> 修改系列的名称
            if (data.name != null && !data.name.isEmpty()) {
                XSSFCell cell = title.getCell(i + 1);
                if (null == cell) {
                    cell = title.createCell(i + 1);
                }
                // 系列名称，不能修改，修改后无法打开 excel
//                cell.setCellValue(data.name);
            }

            // >> 循环每一行的系列数据
            int size = data.value.size();
            for (int j = 0; j < size; j++) {
                // >>> 获取当前行对象(不要第一行标题)
                XSSFRow row = sheet.getRow(j + 1);
                if (row == null) {
                    row = sheet.createRow(j + 1);
                }

                // >>> 获取当前行的维度单元格,并修改维度名
                NameDouble cellValue = data.value.get(j);
                XSSFCell cell = row.getCell(0);
                if (cell == null) {
                    cell = row.createCell(0);
                }
                cell.setCellValue(cellValue.name);

                // >>> 修改当前循环的系列对应的值
                cell = row.getCell(i + 1);
                if (cell == null) {
                    cell = row.createCell(i + 1);
                }
                cell.setCellValue(cellValue.value);
            }

            // > 根据设置的数据删除掉多余的行
            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum > size) {
                for (int idx = lastRowNum; idx > size; idx--) {
                    sheet.removeRow(sheet.getRow(idx));
                }
            }
        }
    }

    /**
     * 更新 chart 的缓存数据
     * @param data          数据
     * @param serTitle      系列的标题缓存
     * @param catDataSource 条目的数据缓存
     * @param numDataSource 数据的缓存
     */
    private void updateChartCatAndNum(SeriesData data, CTSerTx serTitle, CTAxDataSource catDataSource,
            CTNumDataSource numDataSource) {

        // > 更新系列标题
        serTitle.getStrRef().setF(serTitle.getStrRef().getF());
        serTitle.getStrRef().getStrCache().getPtArray(0).setV(data.name);

        // > 也可能是 numRef
        // > 获取cat的数量,val的数量
        long ptCatCnt = catDataSource.getStrRef().getStrCache().getPtCount().getVal();
        long ptNumCnt = numDataSource.getNumRef().getNumCache().getPtCount().getVal();
        int dataSize = data.value.size();
        for (int i = 0; i < dataSize; i++) {
            NameDouble cellValue = data.value.get(i);
            // >> 设置c:cat的c:pt属性
            CTStrVal cat = ptCatCnt > i ? catDataSource.getStrRef().getStrCache().getPtArray(i)
                    : catDataSource.getStrRef().getStrCache().addNewPt();
            cat.setIdx(i);
            cat.setV(cellValue.name);

            // >> 设置c:val的c:pt属性
            CTNumVal val = ptNumCnt > i ? numDataSource.getNumRef().getNumCache().getPtArray(i)
                    : numDataSource.getNumRef().getNumCache().addNewPt();
            val.setIdx(i);
            val.setV(String.format("%.2f", cellValue.value));
        }

        // 更新对应excel的range (<c:f>Sheet1!$B$2:$B$5</c:f>)
        catDataSource.getStrRef().setF(replaceRowEnd(catDataSource.getStrRef().getF(), ptCatCnt, dataSize));
        numDataSource.getNumRef().setF(replaceRowEnd(numDataSource.getNumRef().getF(), ptNumCnt, dataSize));

        // 删除多的c:pt对象
        if (ptNumCnt > dataSize) {
            for (int idx = dataSize; idx < ptNumCnt; idx++) {
                catDataSource.getStrRef().getStrCache().removePt(dataSize);
                numDataSource.getNumRef().getNumCache().removePt(dataSize);
            }
        }
        // 更新个数ptCount
        catDataSource.getStrRef().getStrCache().getPtCount().setVal(dataSize);
        numDataSource.getNumRef().getNumCache().getPtCount().setVal(dataSize);
    }

    /**
     * 替换形如：Sheet1!$A$2:$A$4 的字符
     *
     * @param range
     * @return
     */
    private String replaceRowEnd(String range, long oldSize, long newSize) {
        Pattern pattern = Pattern.compile("(:\\$[A-Z]+\\$)(\\d+)");
        Matcher matcher = pattern.matcher(range);
        if (matcher.find()) {
            long old = Long.parseLong(matcher.group(2));
            return range.replaceAll("(:\\$[A-Z]+\\$)(\\d+)", "$1" + Long.toString(old - oldSize + newSize));
        }
        return range;
    }

    //endregion

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
