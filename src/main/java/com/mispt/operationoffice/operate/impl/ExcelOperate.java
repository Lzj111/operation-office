package com.mispt.operationoffice.operate.impl;

import com.mispt.operationoffice.entity.OperateType;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

/**
 * @Classname ExcelOperate
 * @Description
 * @Date 2022/12/27 17:43
 * @Author by lzj
 */
@Component("excelOperate")
public class ExcelOperate extends BaseOperate {

    @Override
    public void execLogic(String filePath) throws IOException {
        // 1> 构建Excel对象
        FileInputStream fs = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);

        // 2> 循环所有列(getLastRowNum:方法返回的是最后一行的索引，会比总行数少1)
        int rowNum = sheet.getLastRowNum();
        for (int i = 0; i <= rowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            // 3>> 获取列长度
            int cellNum = row.getLastCellNum();
            for (int j = 0; j < cellNum; j++) {
                XSSFCell cell = row.getCell(j);
                if (null == cell) {
                    continue;
                }
                CellType cellType = cell.getCellType();
                // >>>i==0的时候是系列(标题)
                switch (cellType) {
                    case NUMERIC:
                        cell.setCellValue(Math.ceil(Math.random() * 100));
                        break;
                    case STRING:
                        if (!StringUtils.isEmpty(cell.getStringCellValue())) {
                            cell.setCellValue(Math.ceil(Math.random() * 100) + "");
                        }
                        break;
                }
            }
        }

        // > 保存
        FileOutputStream fos = new FileOutputStream(filePath);
        workbook.write(fos);
        fos.flush();
        fos.close();
    }
}
