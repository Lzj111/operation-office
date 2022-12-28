package com.mispt.operationoffice.utils;

import com.mispt.operationoffice.entity.ChartType;
import com.mispt.operationoffice.entity.NameDouble;
import com.mispt.operationoffice.entity.ReplaceValue;
import com.mispt.operationoffice.entity.SeriesData;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
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

    //region 1、完全替换整个数据的方式

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
        );
        return seriesDatas;
    }

    /**
     * 修改图表的数据
     * @param seriesDatas 新的系列数据
     * @param plot 绘图区域对象
     * @param workbook 工作簿对象
     */
    public static void updateChartData(List<SeriesData> seriesDatas, CTPlotArea plot, XSSFWorkbook workbook,
            XDDFChart chart) {
        XSSFSheet sheet = workbook.getSheetAt(0);
        updateChartData(seriesDatas, plot, workbook, sheet, chart);
    }

    /**
     * 修改图表的数据
     * @param seriesDatas 新的系列数据
     * @param plot 绘图区域对象
     * @param workbook 工作簿对象
     * @param sheet 工作表对象
     */
    public static void updateChartData(List<SeriesData> seriesDatas, CTPlotArea plot, XSSFWorkbook workbook,
            XSSFSheet sheet, XDDFChart chart) {
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
                // >> 更新图表Excel的数据
                updateChartExcelV(seriesDatas, workbook, sheet);
                // >> 获取<c:pieChart>的xml对象
                CTPieChart pieChart = plot.getPieChartArray(0);
                // >> 更新chart的缓存数据
                for (int i = 0; i < pieChart.getSerList().size(); i++) {
                    // >>> 获取图表的一个系列对象
                    CTPieSer ser = pieChart.getSerList().get(i);
                    // >>> getTx:系列的标题缓存; getCat:维度的数据缓存; getVal:数据的缓存
                    updateChartCatAndNum(seriesDatas.get(i), ser.getTx(), ser.getCat(), ser.getVal());
                }
            }
            // 3> 折线图
            else if (!plot.getLineChartList().isEmpty()) {
                // >> 更新图表Excel的数据
                updateChartExcelV(seriesDatas, workbook, sheet);
                // >> 获取<c:lineChart>的xml对象
                CTLineChart lineChart = plot.getLineChartArray(0);
                // >> 更新chart的缓存数据
                for (int i = 0; i < lineChart.getSerList().size(); i++) {
                    // >>> 获取图表的一个系列对象
                    CTLineSer ser = lineChart.getSerList().get(i);
                    updateChartCatAndNum(seriesDatas.get(i), ser.getTx(), ser.getCat(), ser.getVal());
                }
            }

            // n> 保存工作簿
            workbook.write(chart.getPackagePart().getOutputStream());
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
        // 判断sheet中是否存在数据行
        if (sheet.getLastRowNum() == 0) {
            return;
        }

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
                cell.setCellValue(data.name);
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
        long ptCatCnt = null != catDataSource ? catDataSource.getStrRef().getStrCache().getPtCount().getVal() : 0;
        long ptNumCnt = null != numDataSource ? numDataSource.getNumRef().getNumCache().getPtCount().getVal() : 0;
        int dataSize = data.value.size();
        for (int i = 0; i < dataSize; i++) {
            NameDouble cellValue = data.value.get(i);
            // >> 设置c:cat的c:pt属性
            if (null != catDataSource) {
                CTStrVal cat = ptCatCnt > i ? catDataSource.getStrRef().getStrCache().getPtArray(i)
                        : catDataSource.getStrRef().getStrCache().addNewPt();
                cat.setIdx(i);
                cat.setV(cellValue.name);
            }

            // >> 设置c:val的c:pt属性
            if (null != numDataSource) {
                CTNumVal val = ptNumCnt > i ? numDataSource.getNumRef().getNumCache().getPtArray(i)
                        : numDataSource.getNumRef().getNumCache().addNewPt();
                val.setIdx(i);
                val.setV(String.format("%.2f", cellValue.value));
            }
        }

        // > 更新对应excel的range (<c:f>Sheet1!$B$2:$B$5</c:f>)
        Optional.ofNullable(catDataSource).ifPresent(cat ->
                cat.getStrRef().setF(replaceRowEnd(cat.getStrRef().getF(), ptCatCnt, dataSize)));
        Optional.ofNullable(numDataSource).ifPresent(val ->
                val.getNumRef().setF(replaceRowEnd(val.getNumRef().getF(), ptNumCnt, dataSize)));

        // > 删除多的c:pt对象
        if (ptNumCnt > dataSize) {
            for (int idx = dataSize; idx < ptNumCnt; idx++) {
                Optional.ofNullable(catDataSource).ifPresent(cat -> cat.getStrRef().getStrCache().removePt(dataSize));
                numDataSource.getNumRef().getNumCache().removePt(dataSize);
            }
        }
        // 更新个数ptCount
        Optional.ofNullable(catDataSource).ifPresent(cat ->
                cat.getStrRef().getStrCache().getPtCount().setVal(dataSize));
        Optional.ofNullable(numDataSource).ifPresent(val ->
                val.getNumRef().getNumCache().getPtCount().setVal(dataSize));
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
    //endregion

    //region 2、通过配置替换数据的方式

    /**
     * 修改图表的数据
     * @param configList 替换配置集合
     * @param plot 绘图区域对象
     * @param workbook 工作簿对象
     * @param chart 图表对象
     */
    public static void updateChartDataConfig(List<ReplaceValue> configList, CTPlotArea plot, XSSFWorkbook workbook,
            XDDFChart chart) {
        XSSFSheet sheet = workbook.getSheetAt(0);
        updateChartDataConfig(configList, plot, workbook, sheet, chart);
    }

    /**
     * 修改图表的数据
     * @param configList 替换配置集合
     * @param plot 绘图区域对象
     * @param workbook 工作簿对象
     * @param sheet 表格对象
     * @param chart 图表对象
     */
    public static void updateChartDataConfig(List<ReplaceValue> configList, CTPlotArea plot, XSSFWorkbook workbook,
            XSSFSheet sheet, XDDFChart chart) {
        try {
            // 1> 柱状图
            if (!plot.getBarChartList().isEmpty()) {
                // >> 获取c:barChart的xml对象
                CTBarChart barChart = plot.getBarChartArray(0);
                // >> 更新图表Excel的数据
                updateChartExcelVConfig(configList, workbook, sheet, barChart, ChartType.CTBarChart);
            }
            // 2> 饼图
            else if (!plot.getPieChartList().isEmpty()) {
                // >> 获取c:pieChart的xml对象
                CTPieChart pieChart = plot.getPieChartArray(0);
                updateChartExcelVConfig(configList, workbook, sheet, pieChart, ChartType.CTPieChart);
            }
            // 3> 折线图
            else if (!plot.getLineChartList().isEmpty()) {
                // >> 获取c:lineChart的xml对象
                CTLineChart lineChart = plot.getLineChartArray(0);
                updateChartExcelVConfig(configList, workbook, sheet, lineChart, ChartType.CTLineChart);
            }

            // n> 保存工作簿
            workbook.write(chart.getPackagePart().getOutputStream());
        } catch (Exception e) {
            System.out.println("ChartUtil.updateChartData异常：" + e.getMessage());
        }
    }

    /**
     * 更新图表的关联 excel，值是纵向的（通过配置更新）
     * @param configList
     * @param workbook
     * @param sheet
     */
    private static void updateChartExcelVConfig(List<ReplaceValue> configList, XSSFWorkbook workbook, XSSFSheet sheet,
            XmlObject xmlObject, ChartType chartType) {
        // 1> 判断sheet中是否存在数据行
        if (sheet.getLastRowNum() == 0) {
            return;
        }

        // 2> 获取行数,并循环行
        // >> getLastRowNum:方法返回的是最后一行的索引，会比总行数少1
        int rowSize = sheet.getLastRowNum();
        for (int i = 0; i <= rowSize; i++) {
            XSSFRow row = sheet.getRow(i);
            if (null == row) {
                row = sheet.createRow(i);
            }

            // >> 获取所有列长度,并循环列
            // >>> getLastCellNum:返回的是最后一列的列数，即等于总列数
            int cellSize = row.getLastCellNum();
            for (int j = 0; j < cellSize; j++) {
                XSSFCell cell = row.getCell(j);
                if (null == cell) {
                    cell = row.createCell(j);
                }
                ReplaceValue replaceValue = getReplaceValue(cell, configList);
                if (null == replaceValue) {
                    continue;
                }
                // >>>> 如果是第一行,表示是系列
                if (i != 0) {
                    // >>>> 修改当前循环的系列对应的配置节点的值(系列名称，不能修改，修改后无法打开 excel)
                    setSheetCellValue(cell, replaceValue);
                }

                // >>>> 更新chart的缓存数据中转方法
                updateChartCatAndNumTransfer(cell, xmlObject, chartType, i, j);
            }
        }
    }

    /**
     * 替换sheet里面单元格数据
     * @param cell
     * @param replaceValue
     */
    private static void setSheetCellValue(XSSFCell cell, ReplaceValue replaceValue) {
        Object value = replaceValue.getValue();
        if (value instanceof Integer) {
            cell.setCellValue(Integer.parseInt(String.valueOf(value)));
        } else if (value instanceof Float) {
            cell.setCellValue(Float.parseFloat(String.valueOf(value)));
        } else if (value instanceof Double) {
            cell.setCellValue(Double.parseDouble(String.valueOf(value)));
        } else if (value instanceof String) {
            cell.setCellValue(String.valueOf(value));
        } else if (value instanceof Date) {
            cell.setCellValue(Date.parse(String.valueOf(value)));
        }
    }

    /**
     * 根据cell对象获取配置节对象
     * @param cell
     * @param configList
     * @return
     */
    private static ReplaceValue getReplaceValue(XSSFCell cell, List<ReplaceValue> configList) {
        // > 只替换编码为字符串的
        CellType cellType = cell.getCellType();
        if (cellType != CellType.STRING) {
            return null;
        }

        // > 获取列值
        String codeContent = cell.getStringCellValue();
        if (null == codeContent) {
            return null;
        }

        // > 循环配置对象
        for (ReplaceValue data : configList) {
            String code = data.getCode();
            if (codeContent.equals(code)) {
                return data;
            }
        }
        return null;
    }

    /**
     * 修改图表缓存数据中转方法
     * @param cell 列对象
     * @param xmlObject 图表对象
     * @param chartType 图表类型
     * @param sheetRowIndex 行索引
     * @param sheetCellIndex 列索引
     */
    private static void updateChartCatAndNumTransfer(XSSFCell cell, XmlObject xmlObject, ChartType chartType,
            int sheetRowIndex, int sheetCellIndex) {
        // 所有<c:ser>下<c:tx>集合(系列的标题缓存)
        List<CTSerTx> serTxs = new ArrayList<>();
        // 所有<c:ser>下<c:cat>集合(维度的缓存)
        List<CTAxDataSource> serCats = new ArrayList<>();
        // 所有<c:ser>下<c:val>集合(数据的缓存)
        List<CTNumDataSource> serVals = new ArrayList<>();

        // > 根据不同图表调用获取不同的配置
        switch (chartType) {
            // > 柱状图
            case CTBarChart:
                CTBarChart barChart = (CTBarChart) xmlObject;
                barChart.getSerList().forEach(ser -> {
                    serTxs.add(ser.getTx());
                    serCats.add(ser.getCat());
                    serVals.add(ser.getVal());
                });
                break;
            // > 饼图
            case CTPieChart:
                CTPieChart pieChart = (CTPieChart) xmlObject;
                pieChart.getSerList().forEach(ser -> {
                    serTxs.add(ser.getTx());
                    serCats.add(ser.getCat());
                    serVals.add(ser.getVal());
                });
                break;
            // 折线图
            case CTLineChart:
                CTLineChart lineChart = (CTLineChart) xmlObject;
                lineChart.getSerList().forEach(ser -> {
                    serTxs.add(ser.getTx());
                    serCats.add(ser.getCat());
                    serVals.add(ser.getVal());
                });
                break;
        }

        // > 更新chart图表里面的缓存数据
        updateChartCatAndNumConfig(cell, serTxs, serCats, serVals, sheetRowIndex, sheetCellIndex);
    }

    /**
     * 更新图表的缓存配置chart.xml的配置数据
     * @param cell sheet的单元格对象
     * @param serTxs 系列标题集合
     * @param serCats cat集合
     * @param serVals val集合
     * @param sheetRowIndex  sheet行索引
     * @param sheetCellIndex sheet列索引
     */
    private static void updateChartCatAndNumConfig(XSSFCell cell, List<CTSerTx> serTxs, List<CTAxDataSource> serCats,
            List<CTNumDataSource> serVals, int sheetRowIndex, int sheetCellIndex) {
        CellType cellType = cell.getCellType();

        // >行索引如果为0的时候,表示更新的系列数据(<c:tx>)
        if (sheetRowIndex == 0) {
            // >>sheetCellIndex为Excel表格的列索引,xml的ser配置里面会少一列维度配置列
            int serIndex = sheetCellIndex - 1;
            // >>获取ser对象
            CTSerTx serTx = serTxs.get(serIndex);
            // >>系列标题应该是文本类型,所以通过文本类型获取表格的对象
            String cellStr = cell.getStringCellValue();
            // >>设置ser的标题缓存
            serTx.getStrRef().setF(serTx.getStrRef().getF());
            serTx.getStrRef().getStrCache().getPtArray(0).setV(cellStr);
            return;
        }

        // >sheetCellIndex为0的时候,表示更新的维度的缓存,需要循环更新ser集合下的所有维度值(<c:cat>)
        else if (sheetCellIndex == 0) {
            // >>循环所有ser下的所有<c:cat>并进行更新维度缓存
            serCats.forEach(serCat -> {
                // >>>cat下ptCount总数(也可能是numRef)
                long catPtCount = serCat.getStrRef().getStrCache().getPtCount().getVal();
                // >>>更新维度缓存时,需要根据sheetRowIndex-1来去掉系列行
                int catCacheIndex = sheetRowIndex - 1;
                // >>>获取pt对象
                CTStrVal catCachePt = catPtCount > catCacheIndex ? serCat.getStrRef().getStrCache().getPtArray(catCacheIndex)
                        : serCat.getStrRef().getStrCache().addNewPt();
                // >>>设置pt对象的索引
                catCachePt.setIdx(catCacheIndex);
                // >>>设置pt对象的v值
                switch (cellType) {
                    case NUMERIC:
                        catCachePt.setV(String.format("%.2f", cell.getNumericCellValue()));
                        break;
                    case STRING:
                        catCachePt.setV(cell.getStringCellValue());
                        break;
                }
            });
        }

        // >其余表示更新具体数据值(<c:val>)
        else {
            // >>更新数据缓存时,需要根据valCellIndex-1来取是哪个系列下的数据
            int valCacheIndex = sheetCellIndex - 1;
            // >>获取第几个系列下的val对象
            CTNumDataSource serVal = serVals.get(valCacheIndex);
            // >>val下ptCount总数(也可能是numRef)
            long valPtCount = serVal.getNumRef().getNumCache().getPtCount().getVal();
            // >>val下的行值获取,需要根据sheetRowIndex-1来剔除第一行系列行
            int valRowIndex = sheetRowIndex - 1;
            CTNumVal valCachePt = valPtCount > valRowIndex ? serVal.getNumRef().getNumCache().getPtArray(valRowIndex)
                    : serVal.getNumRef().getNumCache().addNewPt();
            // >>>设置pt对象的索引
            valCachePt.setIdx(valRowIndex);
            // >>>设置pt对象的v值
            switch (cellType) {
                case NUMERIC:
                    valCachePt.setV(String.format("%.2f", cell.getNumericCellValue()));
                    break;
                case STRING:
                    valCachePt.setV(cell.getStringCellValue());
                    break;
            }
        }
    }
    //endregion

    /**
     * 获取替换对象
     * @param datas 配置数据
     * @param txt 匹配的文本对象
     * @return
     */
    public static ReplaceValue getReplaceValueContainsKey(List<ReplaceValue> datas, String txt) {
        // > 循环配置,找到txt中包含key的对象
        for (int i = 0; i < datas.size(); i++) {
            ReplaceValue data = datas.get(i);
            String code = data.getCode();
            String value = data.getValue().toString();
            if (txt.equals(code)) {
                return data;
            }
        }
        return null;
    }
}
