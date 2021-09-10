package com.ares.excel.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

@ConfigurationProperties(prefix = "ares.excel")
public class AresExcelConfig {

    private String filePath;

    private String idGenerate;

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public String getIdGenerate() {
        return idGenerate;
    }

    public void setIdGenerate(String idGenerate) {
        this.idGenerate = idGenerate;
    }
}
