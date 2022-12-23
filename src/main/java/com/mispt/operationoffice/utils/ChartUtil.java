package com.mispt.operationoffice.utils;

import com.mispt.operationoffice.entity.NameDouble;
import com.mispt.operationoffice.entity.SeriesData;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
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

/**
 * @Classname ChartUtil
 * @Description
 * @Date 2022/12/23 10:37
 * @Author by lzj
 */
public class ChartUtil {

    /**
     * 获取模拟的更新数据
     * @return
     */
    public static List<SeriesData> getSeriesDataList() {
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
                ))
//                new SeriesData("四年级", Arrays.asList(
//                        new NameDouble("2014271班", Math.random() * 100),
//                        new NameDouble("2014272班", Math.random() * 100),
//                        new NameDouble("2014273班", Math.random() * 100),
//                        new NameDouble("2014274班", Math.random() * 100),
//                        new NameDouble("2014275班", Math.random() * 100),
//                        new NameDouble("2014276班", Math.random() * 100)
//                ))
        );
        return seriesDatas;
    }

    /**
     * 修改图表的数据
     * @param seriesDatas 新的系列数据
     * @param plot 绘图区域对象
     * @param workbook 工作簿对象
     */
    public static void updateChartData(List<SeriesData> seriesDatas, CTPlotArea plot, XSSFWorkbook workbook) {
        XSSFSheet sheet = workbook.getSheetAt(0);
        updateChartData(seriesDatas, plot, workbook, sheet);
    }

    /**
     * 修改图表的数据
     * @param seriesDatas 新的系列数据
     * @param plot 绘图区域对象
     * @param workbook 工作簿对象
     * @param sheet 工作表对象
     */
    public static void updateChartData(List<SeriesData> seriesDatas, CTPlotArea plot, XSSFWorkbook workbook, XSSFSheet sheet) {
        try {
            // 1> 柱状图
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
            // 2> 饼图
            else if (!plot.getPieChartList().isEmpty()) {

            }
        } catch (Exception e) {
            System.out.println("ChartUtil.updateChartData异常：" + e.getMessage());
        }
    }

    /**
     * 更新图表的关联 excel，值是纵向的
     * @param seriesDatas
     * @param workbook
     * @param sheet
     */
    private static void updateChartExcelV(List<SeriesData> seriesDatas, XSSFWorkbook workbook, XSSFSheet sheet) {
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
    private static void updateChartCatAndNum(SeriesData data, CTSerTx serTitle, CTAxDataSource catDataSource,
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
    private static String replaceRowEnd(String range, long oldSize, long newSize) {
        Pattern pattern = Pattern.compile("(:\\$[A-Z]+\\$)(\\d+)");
        Matcher matcher = pattern.matcher(range);
        if (matcher.find()) {
            long old = Long.parseLong(matcher.group(2));
            return range.replaceAll("(:\\$[A-Z]+\\$)(\\d+)", "$1" + Long.toString(old - oldSize + newSize));
        }
        return range;
    }

}
