package com.geese.plugin.excel.test;

import com.geese.plugin.excel.mapping.CellMapping;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Created by Administrator on 2017/3/18.
 */
public class TheadLocalTest {

    public static class MyRunnable implements Runnable {
        private ThreadLocal threadLocal = new ThreadLocal() {
            @Override
            protected Object initialValue() {
                return new CellMapping();
            }
        };

        @Override
        public void run() {
            System.out.println(threadLocal.get());;
            try {
                Thread.sleep(10);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            System.out.println(threadLocal.get());
        }
    }

    public static void main(String[] args) {
        MyRunnable sharedRunnable = new MyRunnable();
        Thread thread1 = new Thread(sharedRunnable);
        Thread thread2 = new Thread(sharedRunnable);
        thread1.start();
        thread2.start();
    }


}
