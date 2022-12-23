package com.mispt.operationoffice.operate;

import com.mispt.operationoffice.entity.OperateType;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @Classname IOperate
 * @Description
 * @Date 2022/12/20 15:00
 * @Author by lzj
 */
public interface IOperate {

    void exec(OperateType operateType);

    void execLogic(String filePath) throws IOException;
}
