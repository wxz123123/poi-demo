package com.example.demo.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Description poi service
 * @Author wxz
 * @Date 2019/2/22 15:34
 */
@Service
@Slf4j
public class PoiService {

    /**
     * poi 读excel
     */
    public void readExcel(){
        String filePath = "src/main/resources/test.xlsx";
        //判断是否为excel类型文件
        if(!filePath.endsWith(".xls")&&!filePath.endsWith(".xlsx"))
        {
            log.info("文件不是excel类型");
            return;
        }
        FileInputStream fis =null;
        Workbook wookbook = null;

        //获取一个绝对地址的流
        try {
            fis = new FileInputStream(filePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();

        }
        try {
            //得到工作簿
            wookbook = new XSSFWorkbook(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //得到一个工作表
        Sheet sheet = wookbook.getSheetAt(0);

        //获得表头
        Row rowHead = sheet.getRow(0);
        //判断表头是否正确
//        if(rowHead.getPhysicalNumberOfCells() != 6)
//        {
//            log.info("表头的数量不对!");
//        }
        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        //获得所有数据
        for(int i = 1 ; i <= totalRowNum ; i++)
        {
            //获得第i行对象
            Row row = sheet.getRow(i);

            //获得获得第i行第0列的 String类型对象
            Cell cell = row.getCell(0);
            //设置单元格已文本输出，否则纯数字的单元格，getStringCellValue会报错
            cell.setCellType(Cell.CELL_TYPE_STRING);
            //取出第1列内容
            String id = cell.getStringCellValue();
            //取出第2列内容
            cell = row.getCell(1);
            String name=cell.getStringCellValue();
            log.info("第"+i+"行： "+id+"  "+name);
        }
    }

    /**
     * poi 修改excel
     */
    public void updateExcel(){
        String filePath = "src/main/resources/test.xlsx";
        //判断是否为excel类型文件
        if(!filePath.endsWith(".xls")&&!filePath.endsWith(".xlsx"))
        {
            log.info("文件不是excel类型");
            return;
        }
        FileInputStream fis =null;
        Workbook wookbook = null;

        //获取一个绝对地址的流
        try {
            fis = new FileInputStream(filePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();

        }
        try {
            //得到工作簿
            wookbook = new XSSFWorkbook(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //得到一个工作表
        Sheet sheet = wookbook.getSheetAt(0);

        //获得表头
        Row rowHead = sheet.getRow(0);

        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        //获得所有数据
        for(int i = 1 ; i <= totalRowNum ; i++)
        {
            //获得第i行对象
            Row row = sheet.getRow(i);

            //获得获得第i行第0列的 String类型对象
            Cell cell = row.getCell(0);
            //设置单元格已文本输出，否则纯数字的单元格，getStringCellValue会报错
            cell.setCellType(Cell.CELL_TYPE_STRING);
            // 修改内容
            cell.setCellValue(4);
        }
        //修改完了，一定要保存输出，否则无效
        FileOutputStream fo = null;
        try {
            fo = new FileOutputStream("src/main/resources/test.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            wookbook.write(fo);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
