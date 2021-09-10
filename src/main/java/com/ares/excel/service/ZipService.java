package com.ares.excel.service;

import com.ares.excel.config.AresExcelConfig;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * @author ares
 */
public class ZipService {

    private AresExcelConfig excelConfig;

    public ZipService(AresExcelConfig excelConfig) {
        this.excelConfig = excelConfig;
    }

    /**
     *
     * @param inputFileName 需要打包的文件夹
     * @return
     * @throws Exception
     */
	public String zip(String inputFileName) throws Exception {
        //打包后文件名字
        String zipFileName = excelConfig.getFilePath()+"/"+ UUID.randomUUID().toString()+".zip";
        zip(zipFileName, new File(inputFileName));
        return zipFileName;
    }

    private void zip(String zipFileName, File inputFile) throws Exception {
        ZipOutputStream out = new ZipOutputStream(new FileOutputStream(zipFileName));
        zip(out, inputFile, "");
        out.close();
    }

    private void zip(ZipOutputStream out, File f, String base) throws Exception {
        if (f.isDirectory()) {
           File[] fl = f.listFiles();
           out.putNextEntry(new ZipEntry(base + "/"));
           base = base.length() == 0 ? "" : base + "/";
           for (int i = 0; i < fl.length; i++) {
           zip(out, fl[i], base + fl[i].getName());
         }
        }else {
           out.putNextEntry(new ZipEntry(base));
           FileInputStream in = new FileInputStream(f);
           int b;
           while ( (b = in.read()) != -1) {
            out.write(b);
         }
         in.close();
       }
    }
}