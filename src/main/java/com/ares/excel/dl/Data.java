package com.ares.excel.dl;

import java.util.List;

public class Data<T> {


    private List<T> list;

    private Class<T> clazz;

    public List<T> getList() {
        return list;
    }

    public void setList(List<T> list) {
        this.list = list;
    }

    public Class<T> getClazz() {
        return clazz;
    }

    public void setClazz(Class<T> clazz) {
        this.clazz = clazz;
    }
}
