package com.mispt.operationoffice.operate.impl;

import com.mispt.operationoffice.entity.DataReplace;
import com.mispt.operationoffice.entity.OperateType;
import com.mispt.operationoffice.operate.IOperate;
import java.awt.FileDialog;
import java.awt.Frame;
import java.util.List;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;

/**
 * @Classname BaseOperate
 * @Description
 * @Date 2022/12/20 15:33
 * @Author by lzj
 */
public abstract class BaseOperate implements IOperate {

    private static final int MS_TO_DAYS = 86400000;
    private static final int MS_TO_HOURS = 3600000;
    private static final int MS_TO_MINUTES = 60000;
    private static final int MS_TO_SECONDS = 1000;

    protected final Logger logger = LoggerFactory.getLogger(this.getClass());

    /**
     * 已选文件路径
     */
    private String selFilePath = "";

    /**
     * 数据替换
     */
    @Autowired
    public DataReplace dataReplace;

    /**
     * 毫秒格式化
     *
     * @param mss 毫秒
     * @return java.lang.String
     * @author wlx
     * @date 2021-07-12 14:08
     */
    public static String formatDuring(long mss) {
        long days = mss / MS_TO_DAYS;
        long hours = (mss % MS_TO_DAYS) / MS_TO_HOURS;
        long minutes = (mss % MS_TO_HOURS) / MS_TO_MINUTES;
        long seconds = (mss % MS_TO_MINUTES) / MS_TO_SECONDS;
        return days + " 天 " + hours + " 小时 " + minutes + " 分钟 "
                + seconds + " 秒 ";
    }

    @Override
    public void exec(OperateType operateType) {
        try {
            logger.info("程序执行");
            long start = System.currentTimeMillis();

            // 1> 打开选择框选择要修改的ppt
            FileDialog dialog = new FileDialog(new Frame(), "选择存放位置", FileDialog.LOAD);
            switch (operateType) {
                case PPT:
                    dialog.setFile("*.ppt;*.pptx");
                    break;
                case WORD:
                    dialog.setFile("*.doc;*.docx");
                    break;
                case EXCEL:
                    dialog.setFile("*.xlsx;*.xls");
                    break;
            }
            dialog.setVisible(true);
            String filePath = dialog.getDirectory() + dialog.getFile();
            System.out.println("> 用户选择的文件路径:" + filePath);
            if (!filePath.contains("null")) {
                // 2> 执行处理逻辑
                selFilePath = filePath;
                execLogic(filePath);
            }

            // 3> 计时
            long cost = System.currentTimeMillis() - start;
            String costStr;
            if (cost < MS_TO_SECONDS) {
                costStr = cost + "毫秒";
            } else {
                costStr = formatDuring(cost);
            }
            logger.info("程序执行完成，耗时：" + costStr);
        } catch (Exception e) {
            logger.error("程序执行错误", e);
        }
    }
}
