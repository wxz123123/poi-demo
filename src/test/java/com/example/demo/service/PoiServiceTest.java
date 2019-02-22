package com.example.demo.service;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import static org.junit.Assert.*;


/**
 * @Description poi test
 * @Author wxz
 * @Date 2019/2/22 15:34
 */
@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiServiceTest {
    @Autowired
    private PoiService poiService;
    @Test
    public void readExcel() throws Exception {
        poiService.readExcel();
    }
    @Test
    public void updateExcel() throws Exception {
        poiService.updateExcel();
    }

}