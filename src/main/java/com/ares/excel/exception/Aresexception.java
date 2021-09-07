package com.ares.excel.exception;

public class Aresexception extends RuntimeException {

    private int code;

    public Aresexception(int code,String message){
        super(message);
        this.code = code;
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }
}
