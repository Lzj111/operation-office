package com.mispt.operationoffice.entity;

/**
 * @Classname ReplaceValue
 * @Description
 * @Date 2022/12/26 10:34
 * @Author by lzj
 */
public class ReplaceValue {
    /**
     * 编码
     */
    private String code;
    /**
     * 编码值
     */
    private Object value;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }
}
