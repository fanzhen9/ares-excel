package com.ares.excel.config;

import com.ares.excel.service.ExcelService;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.client.RestTemplate;

import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

@Configuration
@EnableConfigurationProperties(AresExcelConfig.class)
public class EnabelAutoConfiguration {

    @Bean
    @ConditionalOnMissingBean(value = RestTemplate.class)
    public RestTemplate restTemplate(){
        return new RestTemplate();
    }

    @Bean
    @ConditionalOnMissingBean(ThreadPoolExecutor.class)
    public ThreadPoolExecutor executorService(){
        //get cpu core
        int core = Runtime.getRuntime().availableProcessors();
        core = core/2;
        BlockingQueue<Runnable> blockingQueue = new LinkedBlockingQueue();
        return new ThreadPoolExecutor(core, core,
                60L, TimeUnit.SECONDS,blockingQueue
        );
    }

    @Bean
    public ExcelService excelService(AresExcelConfig aresExcelConfig,ThreadPoolExecutor threadPoolExecutor,RestTemplate restTemplate){
        return new ExcelService(aresExcelConfig,threadPoolExecutor,restTemplate);
    }
}
