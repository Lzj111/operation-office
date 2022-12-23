package com.mispt.operationoffice.entity;

import java.util.List;

/**
 * @Description 一个系列的数据
 * @Author lzj
 * @Date 2022/12/21
 **/
public class SeriesData {

    /**
     * value 系列的名字
     */
    public String name;
    public List<NameDouble> value;

    public SeriesData() {
    }

    public SeriesData(List<NameDouble> value) {
        this.value = value;
    }

    public SeriesData(String name, List<NameDouble> value) {
        this.name = name;
        this.value = value;
    }

}
