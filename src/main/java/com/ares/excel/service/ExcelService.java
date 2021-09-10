package com.ares.excel.service;

import com.ares.excel.annotation.Excel;
import com.ares.excel.annotation.ExcelSheet;
import com.ares.excel.config.AresExcelConfig;
import com.ares.excel.dl.Data;
import com.ares.excel.dl.DataL;
import com.ares.excel.service.impl.DefaultDownLoadImageService;
import com.ares.excel.service.impl.UUIDIdGenerate;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.util.StringUtils;
import org.springframework.web.client.RestTemplate;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.Semaphore;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
        writeExcel(list,id);
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
     * 导出excel
     * @param lists
     */
    public void writeExcel(List<Data> lists, String id){
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
            Class clazz = lists.get(i).getClazz();
            XSSFSheet createSheet;
            if(clazz.isAnnotationPresent(ExcelSheet.class)){
                ExcelSheet annotation = (ExcelSheet) clazz.getAnnotation(ExcelSheet.class);
                String name = annotation.name();
                if(!StringUtils.isEmpty(name)){
                    createSheet = workbook.createSheet(name);
                }else{
                    createSheet = workbook.createSheet();
                }
            }else{
                createSheet = workbook.createSheet();
            }


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
            for (int j = 0; j < datalist.size(); j++) {
                XSSFRow dataRow = createSheet.createRow(j+1);
                Field[] dataFields = clazz.getDeclaredFields();
                Object t = datalist.get(j);
                for (Field field : dataFields) {
                    if(field.isAnnotationPresent(Excel.class)){
                        Excel annotation = field.getAnnotation(Excel.class);
                        Integer index = annotation.index();
                        boolean url = annotation.isUrl();
                        Class classType = annotation.fieldClassType();
                        Class loadImageService = annotation.downLoadImageService();

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
            try {
                File file = new File(aresExcelConfig.getFilePath()+"/"+id);
                if(!file.exists()){
                    file.mkdirs();
                }
                FileOutputStream fileOut = new FileOutputStream(aresExcelConfig.getFilePath()+"/"+id+"/"+id+".xlsx");
                workbook.write(fileOut);
                fileOut.close();
            }catch (Exception e){
                e.printStackTrace();

            }
        }
    }

    private Pattern pattern = Pattern.compile("\\{\\.(.*)\\}");

    /**
     * 模板导出
     * @param inputStream
     * @param lists
     */
    public void writeExcel(InputStream inputStream,List<Data> lists,String id){
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            int sheetNum = lists.size();
            for (int i = 0; i < sheetNum; i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                //get all field
                XSSFRow row = sheet.getRow(1);
                int physicalNumberOfCells = row.getPhysicalNumberOfCells();
                List<DataL> list = new ArrayList<>();
                for (int j = 0; j < physicalNumberOfCells; j++) {
                    String key = row.getCell(j).getStringCellValue();
                    Matcher matcher = pattern.matcher(key);
                    if (matcher.find()) {
                        list.add(new DataL(j,matcher.group(1)));
                    }
                }
                //getData
                List dataList = lists.get(i).getList();
                Class dataType = lists.get(i).getClazz();
                for (int k=0 ; k<dataList.size() ; k++) {
                    XSSFRow rowM = sheet.getRow(k+1);
                    if(rowM == null){
                        rowM = sheet.createRow(k+1);
                    }
                    for (DataL dataL : list) {
                        Object obj = dataList.get(k);
                        String keyName = dataL.getKeyName();
                        String str1 = keyName.substring(0, 1);
                        String str2 = keyName.substring(1);
                        String methodName = "get" + str1.toUpperCase() + str2;
                        Method method = dataType.getMethod(methodName);
                        Object object = method.invoke(obj);
                        Class classType = dataType.getDeclaredField(keyName).getType();
                        XSSFCell cell = rowM.getCell(dataL.getIndex());
                        if(cell == null){
                            cell = rowM.createCell(dataL.getIndex());
                        }
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
                    }
                }
            }
            File file = new File(aresExcelConfig.getFilePath()+"/"+id);
            if(!file.exists()){
                file.mkdirs();
            }
            FileOutputStream fileOut = new FileOutputStream(aresExcelConfig.getFilePath()+"/"+id+"/"+id+".xlsx");
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e){
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * @param is
     * @param clazz
     * @param readline
     * @param <T>
     * @return
     */
    public <T> List<T> readExcel(InputStream is,int sheetIndex,Class<T> clazz,int readline) throws IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException, IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        XSSFSheet sheetAt = workbook.getSheetAt(sheetIndex);
        Integer rows = sheetAt.getLastRowNum();

        List<T> result = new ArrayList<T>();
        for (; readline <= rows; readline++) {
            XSSFRow row = sheetAt.getRow(readline);
            T t = clazz.newInstance();
            Field[] declaredFields = clazz.getDeclaredFields();
            for (Field field : declaredFields) {
                if(field.isAnnotationPresent(Excel.class)){
                    Excel annotation = field.getAnnotation(Excel.class);
                    int columnIndex = annotation.index();
                    Class aClass = annotation.fieldClassType();
                    String methodName = field.getName();
                    String str1 = methodName.substring(0, 1);
                    String str2 = methodName.substring(1);
                    String methodSet = "set" + str1.toUpperCase() + str2;
                    Method method = clazz.getMethod(methodSet,aClass);
                    XSSFCell cell = row.getCell(columnIndex);
                    if(String.class == aClass && cell != null){
                        String value = cell.getStringCellValue();
                        method.invoke(t,value);
                    }
                    if((Integer.class == aClass||Double.class == aClass||Float.class == aClass) && cell != null){
                        double value = cell.getNumericCellValue();
                        method.invoke(t,aClass.cast(value));
                    }
                    if(Boolean.class == aClass && cell != null){
                        boolean value = cell.getBooleanCellValue();
                        method.invoke(t,aClass.cast(value));
                    }
                    if(Date.class == aClass && cell != null){
                        Date value = cell.getDateCellValue();
                        method.invoke(t,aClass.cast(value));
                    }
                }
            }
            result.add(t);
        }
        return result;
    }
}
