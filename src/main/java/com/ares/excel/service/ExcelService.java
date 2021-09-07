package com.ares.excel.service;

import com.ares.excel.annotation.Excel;
import com.ares.excel.config.AresExcelConfig;
import com.ares.excel.exception.Aresexception;
import com.ares.excel.service.impl.UUIDIdGenerate;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.util.StringUtils;
import org.springframework.web.client.RestTemplate;

import java.lang.reflect.Field;
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
     * @param lists
     * @param clazzs
     * @param <T>
     */
    public <T> void writeExcel(List<List<T>> lists,List<Class> clazzs){
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
        writeExcel(lists,clazzs,id,null);
    }

    /**
     *
     * @param lists
     * @param clazzs
     * @param <T>
     */
    public <T> void writeExcel(List<List<T>> lists,List<Class> clazzs, String id){
        writeExcel(lists,clazzs,id,null);
    }

    /**
     *
     * @param lists
     * @param clazzs
     * @param <T>
     */
    public <T> void writeExcel(List<List<T>> lists,List<Class> clazzs, String id,DownLoadImageService downLoadImageService){
        if(lists.size()!=clazzs.size()){
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
            //创建第一行
            Class clazz = clazzs.get(i);
            XSSFRow row = createSheet.createRow(0);
            Field[] fields = clazz.getDeclaredFields();
            Semaphore semaphore = new Semaphore(20);
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
            List<T> list = lists.get(i);
            final CountDownLatch countDownLatch = new CountDownLatch(list.size() * urlCount);

        }
    }
}
