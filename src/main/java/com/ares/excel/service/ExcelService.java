package com.ares.excel.service;

import com.ares.excel.annotation.Excel;
import com.ares.excel.config.AresExcelConfig;
import com.ares.excel.dl.Data;
import com.ares.excel.exception.Aresexception;
import com.ares.excel.service.impl.DefaultDownLoadImageService;
import com.ares.excel.service.impl.UUIDIdGenerate;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.util.StringUtils;
import org.springframework.web.client.RestTemplate;

import java.io.File;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.Semaphore;
import java.util.concurrent.ThreadPoolExecutor;

public class ExcelService {

    private AresExcelConfig aresExcelConfig;

    private ThreadPoolExecutor threadPoolExecutor;

    private RestTemplate restTemplate;

    private UUIDIdGenerate uuidIdGenerate;

    public ExcelService(AresExcelConfig aresExcelConfig,ThreadPoolExecutor threadPoolExecutor,RestTemplate restTemplate) {
        this.aresExcelConfig = aresExcelConfig;
        this.threadPoolExecutor = threadPoolExecutor;
        this.restTemplate = restTemplate;
        uuidIdGenerate = new UUIDIdGenerate();
    }
    //
    /**
     * 导出excle
     * @param list
     */
    public void writeExcel(List<Data> list){
        String id = getSimpleId();
        writeExcel(list,id,null);
    }

    private String getSimpleId() {
        String idGenerate = aresExcelConfig.getIdGenerate();
        String id = "";
        if(StringUtils.isEmpty(idGenerate)){
            id = uuidIdGenerate.getName();
        }else{
            try {
                Class<?> clazz = Class.forName(idGenerate);
                Object o = clazz.newInstance();
                IdGenerate g = (IdGenerate) o;
                id = g.getName();
            }catch (ClassNotFoundException e){
                e.printStackTrace();
            } catch (InstantiationException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
        return id;
    }

    /**
     * 导出excle
     * @param list
     */
    public <T> void writeExcel(List<Data> list,String id){
        writeExcel(list,id,null);
    }

    /**
     * 导出excel
     * @param lists
     */
    public void writeExcel(List<Data> lists, String id,DownLoadImageService downLoadImageService){
        if(lists.size()!=0){
            //TODO 最后设计枚举报错内容
            //throw new Aresexception();
        }

        //构建excel文档
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFCellStyle linkStyle = workbook.createCellStyle();
        XSSFFont cellFont = workbook.createFont();
        cellFont.setUnderline((byte) 1);
        cellFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        linkStyle.setFont(cellFont);

        for (int i = 0; i < lists.size(); i++) {
            //开始拼excel

            //构建sheet
            XSSFSheet createSheet = workbook.createSheet();
            Class clazz = lists.get(i).getClazz();
            Field[] fields = clazz.getDeclaredFields();
            //创建第一行
            XSSFRow row = createSheet.createRow(0);
            int urlCount = 0;
            for (Field field : fields) {
                if(field.isAnnotationPresent(Excel.class)){
                    Excel annotation = field.getAnnotation(Excel.class);
                    Integer index = annotation.index();
                    String name = annotation.name();
                    XSSFCell cell = row.createCell(index);
                    boolean url = annotation.isUrl();
                    if(url){
                        urlCount ++;
                    }
                    cell.setCellValue(name);
                }
            }
            List datalist = lists.get(i).getList();

            Semaphore semaphore = new Semaphore(20);
            final CountDownLatch countDownLatch = new CountDownLatch(datalist.size() * urlCount);
            for (int j = 0; i < datalist.size(); j++) {
                XSSFRow dataRow = createSheet.createRow(i+1);
                Field[] dataFields = clazz.getDeclaredFields();
                Object t = datalist.get(i);
                for (Field field : dataFields) {
                    if(field.isAnnotationPresent(Excel.class)){
                        Excel annotation = field.getAnnotation(Excel.class);
                        Integer index = annotation.index();
                        boolean url = annotation.isUrl();
                        Class classType = annotation.fieldClassType();
                        Class loadImageService = annotation.DownLoadImageService();

                        XSSFCell cell = dataRow.createCell(index);
                        try {
                            String methodName = field.getName();
                            String str1 = methodName.substring(0, 1);
                            String str2 = methodName.substring(1);
                            String methodGet = "get" + str1.toUpperCase() + str2;
                            Method method = clazz.getMethod(methodGet);
                            Object object = method.invoke(t);
                            if (object == null) {
                                classType = String.class;
                                object = "";
                            }
                            if(String.class == classType){
                                cell.setCellValue(String.valueOf(object));
                            }
                            if(Double.class == classType||Integer.class == classType){
                                cell.setCellValue(Double.valueOf(String.valueOf(object)));
                            }
                            if(Date.class == classType){
                                cell.setCellValue((Date)object);
                            }
                            if(Boolean.class == classType){
                                cell.setCellValue(Boolean.valueOf(String.valueOf(object)));
                            }
                            if(url){
                                File file = new File(aresExcelConfig.getFilePath()+"/"+id+"/pic");
                                if(!file.exists()){
                                    file.mkdirs();
                                }
                                String fileName = getSimpleId()+".jpg";
                                semaphore.acquire();
                                //开始图片下载
                                if(loadImageService==DefaultDownLoadImageService.class){
                                    threadPoolExecutor.execute(new DefaultDownLoadImageService(aresExcelConfig.getFilePath()+"/"+id+"/pic",fileName,String.valueOf(object),countDownLatch,restTemplate));
                                }else{
                                    Constructor constructor = loadImageService.getConstructor(String.class, String.class, String.class, CountDownLatch.class, RestTemplate.class);
                                    AbstractDownLoadImageService o = (AbstractDownLoadImageService)constructor.newInstance(aresExcelConfig.getFilePath() + "/" + id + "/pic", fileName, String.valueOf(object), countDownLatch, restTemplate);
                                    threadPoolExecutor.execute(o);
                                }
                                semaphore.release();
                                CreationHelper createHelper = workbook.getCreationHelper();
                                XSSFHyperlink hyperlink = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.FILE);
                                hyperlink.setAddress("./pic/"+fileName);
                                cell.setHyperlink(hyperlink);
                                cell.setCellValue("【点击打开】");
                                cell.setCellStyle(linkStyle);
                            }
                        } catch (NoSuchMethodException e) {
                            e.printStackTrace();
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        } catch (InvocationTargetException e) {
                            e.printStackTrace();
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        } catch (InstantiationException e) {
                            e.printStackTrace();
                        }

                    }
                }
            }
            try {
                countDownLatch.await();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }
}
