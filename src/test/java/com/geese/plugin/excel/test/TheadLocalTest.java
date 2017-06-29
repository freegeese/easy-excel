package com.geese.plugin.excel.test;

import com.geese.plugin.excel.ExcelContext;
import com.geese.plugin.excel.ExcelTemplate;
import com.geese.plugin.excel.mapping.CellMapping;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Created by Administrator on 2017/3/18.
 */
public class TheadLocalTest {


    public static void main(String[] args) throws InterruptedException {
        final Service service = new Service();
        for (int i = 0; i < 5; i++) {
            new Thread() {
                @Override
                public void run() {
                    service.doService();
                }
            }.start();
        }
        Thread.sleep(1000L);
        System.out.println(ExcelContext.get());
    }

    public static class Service {
        public synchronized void doService() {
            Map map = ExcelContext.get();
            System.out.println(map.get("a"));
            map.put("a", Thread.currentThread().getName() + Math.random() * 1000);
            System.out.println(ExcelContext.get().get("a"));
        }

    }

}
