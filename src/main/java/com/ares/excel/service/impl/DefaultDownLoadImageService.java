package com.ares.excel.service.impl;

import com.ares.excel.service.AbstractDownLoadImageService;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.client.RestTemplate;

import java.io.File;
import java.io.FileOutputStream;
import java.util.concurrent.CountDownLatch;

public class DefaultDownLoadImageService extends AbstractDownLoadImageService {

    public DefaultDownLoadImageService(String path, String fileName, String url, CountDownLatch countDownLatch, RestTemplate restTemplate) {
        super(path,fileName,url,countDownLatch,restTemplate);
    }

    @Override
    public void downLoadImage() {
        try {
            ResponseEntity<byte[]> responseEntity = restTemplate.exchange(url, HttpMethod.GET, null, byte[].class);
            HttpStatus statusCode = responseEntity.getStatusCode();
            if (statusCode == HttpStatus.NOT_FOUND) {
                return;
            }
            //获取entity中的数据
            byte[] body = responseEntity.getBody();
            //创建输出流  输出到本地
            FileOutputStream fileOutputStream = new FileOutputStream(new File(path + "/" + fileName));
            fileOutputStream.write(body);
            fileOutputStream.flush();
            //关闭流
            fileOutputStream.close();
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            countDownLatch.countDown();
        }
    }


}
