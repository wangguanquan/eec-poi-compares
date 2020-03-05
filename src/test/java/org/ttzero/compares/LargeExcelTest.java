/*
 * Copyright (c) 2017-2020, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


package org.ttzero.compares;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.junit.Before;
import org.junit.FixMethodOrder;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Sheet;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static org.junit.runners.MethodSorters.NAME_ASCENDING;
import static org.ttzero.compares.BaseExcelTest.random;

/**
 * @author guanquan.wang at 2020-02-28 14:22
 */
@FixMethodOrder(NAME_ASCENDING)
public class LargeExcelTest {
    private static final Logger LOGGER = LoggerFactory.getLogger(LargeExcelTest.class);

    public static Path defaultTestPath = Paths.get("out/excel/");
    private static File template07;
    private int i, loop = 1;

    @Before public void before() throws IOException {
        if (!Files.exists(defaultTestPath)) {
            Files.createDirectories(defaultTestPath);
        }
        template07 = new File("./src/test/resources/large/fill.xlsx");
    }

    // -----------------1w----------------
    @Test public void testEasy1w() {
        loop = 10;
        easyWrite("easy 1w");
    }


    @Test public void testEec1w() throws IOException {
        loop = 10;
        eecWrite("eec 1w");
    }

    @Test public void testEasy1wr() {
        easyRead("easy 1w");
    }

    @Test public void testEec1wr() {
        eecRead("eec 1w");
    }

    // -----------------10w----------------
    @Test public void testEasy10w() {
        loop = 100;
        easyWrite("easy 10w");
    }


    @Test public void testEec10w() throws IOException {
        loop = 100;
        eecWrite("eec 10w");
    }

    @Test public void testEasy10wr() {
        easyRead("easy 10w");
    }

    @Test public void testEec10wr() {
        eecRead("eec 10w");
    }

    // -----------------50w----------------
    @Test public void testEasy50w() {
        loop = 500;
        easyWrite("easy 50w");
    }


    @Test public void testEec50w() throws IOException {
        loop = 500;
        eecWrite("eec 50w");
    }

    @Test public void testEasy50wr() {
        easyRead("easy 50w");
    }

    @Test public void testEec50wr() {
        eecRead("eec 50w");
    }

    // -----------------100w----------------
    @Test public void testEasy100w() {
        loop = 1000;
        easyWrite("easy 100w");
    }


    @Test public void testEec100w() throws IOException {
        loop = 1000;
        eecWrite("eec 100w");
    }

    @Test public void testEasy100wr() {
        easyRead("easy 100w");
    }

    @Test public void testEec100wr() {
        eecRead("eec 100w");
    }

    private void easyWrite(String name) {
        LOGGER.info("Easy-excel start to write...");
        long start = System.currentTimeMillis();
        ExcelWriter excelWriter = EasyExcel.write(defaultTestPath.resolve(name + ".xlsx").toFile()).withTemplate(template07).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        for (int j = 0; j < loop; j++) {
            excelWriter.fill(data(), writeSheet);
            LOGGER.info("{} fill success.", j);
        }
        excelWriter.finish();
        LOGGER.info("Easy-excel write finished. used: {}", System.currentTimeMillis() - start);
    }

    private void eecWrite(String name) throws IOException {
        LOGGER.info("EEC start to write...");
        long start = System.currentTimeMillis();
        new Workbook().addSheet(new ListSheet<LargeData>() {
            int n = 0;
            public List<LargeData> more() {
                LOGGER.info("{} fill success.", n);
                return n++ < loop ? data() : null;
            }
        }).writeTo(defaultTestPath.resolve(name + ".xlsx"));
        LOGGER.info("EEC write finished. used: {}", System.currentTimeMillis() - start);
    }

    private void easyRead(String name) {
        LOGGER.info("Easy-excel start to read...");
        long start = System.currentTimeMillis();
        EasyExcel.read(defaultTestPath.resolve(name + ".xlsx").toFile(), LargeData.class,
            new LargeDataListener()).headRowNumber(1).sheet().doRead();
        LOGGER.info("Easy-excel read finished. used: {}", System.currentTimeMillis() - start);
    }

    private void eecRead(String name) {
        LOGGER.info("EEC start to read...");
        long start = System.currentTimeMillis();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(name + ".xlsx"))) {
            long n = reader.sheets().flatMap(Sheet::dataRows).map(row -> row.too(LargeData.class)).count();
            LOGGER.info("Data rows: {}", n);
        } catch (IOException e) {
            e.printStackTrace();
        }
        LOGGER.info("EEC read finished. used: {}", System.currentTimeMillis() - start);
    }


    // 以下代码由easyexcel测试代码`com.alibaba.easyexcel.test.core.large.LargeDataTest#data`复制而来
    // 在原来的列上加了4个基础类型测试
    private List<LargeData> data() {
        List<LargeData> list = new ArrayList<>();
        int size = i + 1000;
        for (; i < size; i++) {
            LargeData largeData = new LargeData();
            list.add(largeData);
            largeData.setNv(random.nextInt());
            largeData.setLv(random.nextLong());
            largeData.setDv(random.nextDouble());
            largeData.setAv(new Date(System.currentTimeMillis() - random.nextInt(9999)));
            largeData.setStr1("str1-" + i);
            largeData.setStr2("str2-" + i);
            largeData.setStr3("str3-" + i);
            largeData.setStr4("str4-" + i);
            largeData.setStr5("str5-" + i);
            largeData.setStr6("str6-" + i);
            largeData.setStr7("str7-" + i);
            largeData.setStr8("str8-" + i);
            largeData.setStr9("str9-" + i);
            largeData.setStr10("str10-" + i);
            largeData.setStr11("str11-" + i);
            largeData.setStr12("str12-" + i);
            largeData.setStr13("str13-" + i);
            largeData.setStr14("str14-" + i);
            largeData.setStr15("str15-" + i);
            largeData.setStr16("str16-" + i);
            largeData.setStr17("str17-" + i);
            largeData.setStr18("str18-" + i);
            largeData.setStr19("str19-" + i);
            largeData.setStr20("str20-" + i);
            largeData.setStr21("str21-" + i);
            largeData.setStr22("str22-" + i);
            largeData.setStr23("str23-" + i);
            largeData.setStr24("str24-" + i);
            largeData.setStr25("str25-" + i);
        }
        return list;
    }
}
