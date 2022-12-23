package com.mispt.operationoffice.operate;

import com.mispt.operationoffice.entity.OperateType;
import com.mispt.operationoffice.utils.BeanUtil;
import java.util.Objects;

/**
 * @Classname OperateTemplate
 * @Description
 * @Date 2022/12/20 15:05
 * @Author by lzj
 */
public class OperateTemplate {

    public static IOperate getOperate(OperateType operateType) {
        Objects.requireNonNull(operateType);
        String beanName = null;
        switch (operateType) {
            case PPT:
                beanName = "pptOperate";
                break;
            case WORD:
                beanName = "worldOperate";
                break;
            default:
                beanName = "operate";
                break;
        }
        return BeanUtil.getInstanceByBeanName(beanName);
    }
}
