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
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import org.junit.Test;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.WaterMark;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Sheet;

import java.awt.*;
import java.io.IOException;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Random;

import static org.ttzero.compares.LargeExcelTest.defaultTestPath;

/**
 * @author guanquan.wang at 2020-03-05 15:42
 */
public class BaseExcelTest {

    private Converter charConverter = new Converter<Character>() {
        @Override
        public Class supportJavaTypeKey() {
            return Character.class;
        }

        @Override
        public CellDataTypeEnum supportExcelTypeKey() {
            return CellDataTypeEnum.STRING;
        }

        @Override
        public Character convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
            String s;
            return (s = cellData.getStringValue()) != null && s.length() >= 1 ? s.charAt(0) : '\0';
        }

        @Override
        public CellData convertToExcelData(Character value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
            return new CellData(new String(new char[] {value}));
        }
    };

    private Converter timestampConverter = new Converter<Timestamp>() {
        @Override
        public Class supportJavaTypeKey() {
            return Timestamp.class;
        }

        @Override
        public CellDataTypeEnum supportExcelTypeKey() {
            return CellDataTypeEnum.NUMBER;
        }

        @Override
        public Timestamp convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
            return Timestamp.valueOf(cellData.getStringValue());
        }

        @Override
        public CellData convertToExcelData(Timestamp value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
            return new CellData(value.toString());
        }
    };

    private Converter timeConverter = new Converter<Time>() {
        @Override
        public Class supportJavaTypeKey() {
            return Time.class;
        }

        @Override
        public CellDataTypeEnum supportExcelTypeKey() {
            return CellDataTypeEnum.STRING;
        }

        @Override
        public Time convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
            return Time.valueOf(cellData.getStringValue());
        }

        @Override
        public CellData convertToExcelData(Time value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
            return new CellData(value.toString());
        }
    };

    @Test public void test1() {
        EasyExcel.write(defaultTestPath.resolve("testEasyExcelAllType.xlsx").toString())
            .registerConverter(charConverter)
            .registerConverter(timestampConverter)
            .registerConverter(timeConverter)
            .sheet().doWrite(AllType.randomTestData());
    }

    @Test public void test2() throws IOException {
        new Workbook()
            .addSheet(AllType.randomTestData())
            .writeTo(defaultTestPath.resolve("testEECAllType.xlsx"));
    }

    @Test public void test3() {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("testEECAllType.xlsx"))) {
            reader.sheets().flatMap(Sheet::rows).forEach(System.out::println);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void test4() {
        EasyExcel.read(defaultTestPath.resolve("testEasyExcelAllType.xlsx").toFile(), simpleListener)
            .registerConverter(charConverter)
            .registerConverter(timestampConverter)
            .registerConverter(timeConverter)
            .headRowNumber(1).sheet().doRead();
    }


    @Test public void test5() {
        EasyExcel.write(defaultTestPath.resolve("test5.xlsx").toString(), Check.class).sheet().doWrite(checks());
    }

    @Test public void test6() throws IOException {
        new Workbook("test6").addSheet(new ListSheet<>(checks())).writeTo(defaultTestPath);
    }

    @Test public void test7() {
        ExcelWriter excelWriter = EasyExcel.write(defaultTestPath.resolve("test7.xlsx").toString()).build();
        excelWriter.write(checks(), EasyExcel.writerSheet("帐单表").build());
        excelWriter.write(customers(), EasyExcel.writerSheet("客户表").build());
        excelWriter.write(c2CS(), EasyExcel.writerSheet("用户客户关系表").build());
        excelWriter.finish();
    }

    @Test public void test8() throws IOException {
        new Workbook("test8")
            .addSheet(new ListSheet<>("帐单表", checks()))
            .addSheet(new ListSheet<>("客户表", customers()))
            .addSheet(new ListSheet<>("用户客户关系表", c2CS()))
            .writeTo(defaultTestPath);
    }

    @Test public void testStyleConversion() throws IOException {
        new Workbook("testStyleConversion") // 文件名
            .setCreator("奈留·智库") // 作者
            .setCompany("Copyright (c) 2020") // 公司名
            .setWaterMark(WaterMark.of("Secret")) // 水印
            .setAutoSize(true) // 自动计算列宽
            .addSheet(new ListSheet<>("期末成绩", Student.randomTestData(20)
                , new org.ttzero.excel.entity.Sheet.Column("学号", "id", int.class)
                , new org.ttzero.excel.entity.Sheet.Column("姓名", "name", String.class)
                , new org.ttzero.excel.entity.Sheet.Column("成绩", "score", int.class)
                // 低于60分显示`不及格`
                .setProcessor(n -> n < 60 ? "不及格" : n)
                // 低于60分单元格标红
                .setStyleProcessor((o, style, sst) -> {
                    if ((int)o < 60) {
                        style = Styles.clearFill(style) | sst.addFill(new Fill(Color.red));
                    }
                    return style;
                })
            )
        )
        .writeTo(defaultTestPath);
    }

    static ReadListener<Map<String, Object>> simpleListener = new AnalysisEventListener<Map<String, Object>> () {
        @Override
        public void invoke(Map<String, Object> data, AnalysisContext context) {
            System.out.println(data);
        }

        @Override
        public void doAfterAllAnalysed(AnalysisContext context) { }
    };

    @Test public void test9() {
        com.alibaba.excel.ExcelReader excelReader = EasyExcel.read(defaultTestPath.resolve("test7.xlsx").toFile(), simpleListener).headRowNumber(0).build();
        List<ReadSheet> sheets = excelReader.excelExecutor().sheetList();
        sheets.forEach(sheet -> {
            System.out.println("----------" + sheet.getSheetName() + "-----------");
            excelReader.read(sheet);
        });
    }

    @Test public void test10() {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test8.xlsx"))) {
            reader.sheets().flatMap(sheet -> {
                System.out.println("----------" + sheet.getName() + "-----------");
                return sheet.rows();
            }).forEach(System.out::println);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

































    private java.util.List<Check> checks() {
        return Arrays.asList(new Check(1, 100.8), new Check(2, 34.2), new Check(3, 983));
    }

    private java.util.List<Customer> customers() {
        return Arrays.asList(new Customer(1001, "张三"), new Customer(1002, "李四"));
    }

    private List<C2C> c2CS() {
        return Arrays.asList(new C2C(1, 1001), new C2C(2, 1002), new C2C(3, 1002));
    }


    public static Random random = new Random();

    public static char[] charArray = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890".toCharArray();
    private static char[] cache = new char[32];
    public static String getRandomString() {
        int n = random.nextInt(cache.length) + 1, size = charArray.length;
        for (int i = 0; i < n; i++) {
            cache[i] = charArray[random.nextInt(size)];
        }
        return new String(cache, 0, n);
    }
}
