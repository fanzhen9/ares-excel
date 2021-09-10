package com.ares.excel.dl;


public class DataL {

    private int index;

    private String keyName;

    public DataL() {
    }

    public DataL(int index, String keyName) {
        this.index = index;
        this.keyName = keyName;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getKeyName() {
        return keyName;
    }

    public void setKeyName(String keyName) {
        this.keyName = keyName;
    }
}
