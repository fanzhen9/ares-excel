package com.ares.excel.service.impl;

import com.ares.excel.service.IdGenerate;

import java.util.UUID;

public class UUIDIdGenerate implements IdGenerate {


    @Override
    public String getName() {
        UUID uuid = UUID.randomUUID();
        return uuid.toString().replaceAll("-","");
    }
}
