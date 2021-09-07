package com.ares.excel.service;

import org.springframework.web.client.RestTemplate;

import java.util.concurrent.CountDownLatch;

public abstract class AbstractDownLoadImageService implements DownLoadImageService,Runnable{

    protected String path;

    protected String fileName;

    protected String url;

    protected CountDownLatch countDownLatch;

    protected RestTemplate restTemplate;

    public AbstractDownLoadImageService(String path, String fileName, String url, CountDownLatch countDownLatch, RestTemplate restTemplate) {
        this.path = path;
        this.fileName = fileName;
        this.url = url;
        this.countDownLatch = countDownLatch;
        this.restTemplate = restTemplate;
    }

    @Override
    public void run() {
        downLoadImage();
    }
}
