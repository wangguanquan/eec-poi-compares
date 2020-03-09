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
    private final Logger LOGGER = LoggerFactory.getLogger(getClass());

    public static Path defaultTestPath = Paths.get("out/excel/");
    private static File template07;
    private int i, loop = 1;

    @Before public void before() throws IOException {
        if (!Files.exists(defaultTestPath)) {
            Files.createDirectories(defaultTestPath);
        }
        template07 = new File("./src/test/resources/large/fill.xlsx");
    }

    // 先空测写1w ~ 100w的时间
    // -----------------1w----------------
    @Test public void test1w() {
        loop = 10;
        emptyLoop();
    }

    // -----------------5w----------------
    @Test public void test5w() {
        loop = 50;
        emptyLoop();
    }

    // -----------------10w----------------
    @Test public void test10w() {
        loop = 100;
        emptyLoop();
    }

    // -----------------50w----------------
    @Test public void test50w() {
        loop = 500;
        emptyLoop();
    }

    // -----------------100w----------------
    @Test public void test100w() throws InterruptedException {
        Thread.sleep(20_000L);
        loop = 1000;
        emptyLoop();
    }

    private void emptyLoop() {
        for (int j = 0; j < loop; j++, data());
    }

    // 以下进行excel读写测试
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

    // -----------------5w----------------
    @Test public void testEasy5w() {
        loop = 50;
        easyWrite("easy 5w");
    }


    @Test public void testEec5w() throws IOException {
        loop = 50;
        eecWrite("eec 5w");
    }

    @Test public void testEasy5wr() {
        easyRead("easy 5w");
    }

    @Test public void testEec5wr() {
        eecRead("eec 5w");
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

    // ------------------JDBC似读文件------------------
//    @Test public void testEec1wrJdbc() {
//        eecRead0("eec 1w");
//    }
//
//    @Test public void testEec5wrJdbc() {
//        eecRead0("eec 5w");
//
//    }
//
//    @Test public void testEec10wrJdbc() {
//        eecRead0("eec 10w");
//
//    }
//
//    @Test public void testEec50wrJdbc() {
//        eecRead0("eec 50w");
//
//    }
//
//    @Test public void testEec100wrJdbc() {
//        eecRead0("eec 100w");
//    }
//
//    //------------------SharedString方式------------------
//    //-----------------1w----------------
//    @Test public void testEecShared1w() throws IOException {
//        loop = 10;
//        eecWriteShared("eec shared 1w");
//    }
//
//    @Test public void testEecShared1wr() {
//        eecRead("eec shared 1w");
//    }
//
//    @Test public void testEsyShared1wr() {
//        easyRead("eec shared 1w");
//    }
//
//    //-----------------5w----------------
//    @Test public void testEecShared5w() throws IOException {
//        loop = 50;
//        eecWriteShared("eec shared 5w");
//    }
//
//    @Test public void testEecShared5wr() {
//        eecRead("eec shared 5w");
//    }
//
//    @Test public void testEsyShared5wr() {
//        easyRead("eec shared 5w");
//    }
//
//    //-----------------10w----------------
//    @Test public void testEecShared10w() throws IOException {
//        loop = 100;
//        eecWriteShared("eec shared 10w");
//    }
//
//    @Test public void testEecShared10wr() {
//        eecRead("eec shared 10w");
//    }
//
//    @Test public void testEsyShared10wr() {
//        easyRead("eec shared 10w");
//    }
//
//    //-----------------50w----------------
//    @Test public void testEecShared50w() throws IOException {
//        loop = 500;
//        eecWriteShared("eec shared 50w");
//    }
//
//    @Test public void testEecShared50wr() {
//        eecRead("eec shared 50w");
//    }
//
//    @Test public void testEsyShared50wr() {
//        easyRead("eec shared 50w");
//    }
//
//    //-----------------100w----------------
//    @Test public void testEecShared100w() throws IOException {
//        loop = 1000;
//        eecWriteShared("eec shared 100w");
//    }
//
//    @Test public void testEecShared100wr() {
//        eecRead("eec shared 100w");
//    }
//
//    @Test public void testEsyShared100wr() {
//        easyRead("eec shared 100w");
//    }

    //--------------------------------------

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
            long n = reader.sheets().flatMap(sheet -> {
                LOGGER.info("Worksheet [{}] dimension: {}", sheet.getName(), sheet.getDimension());
                return sheet.dataRows();
            }).map(row -> {
                if (row.getRowNumber() % 100_000 == 0) {
                    LOGGER.info("Reading {} rows", row.getRowNumber());
                }
                return row.too(LargeData.class);
            }).count();
            LOGGER.info("Data rows: {}", n);
        } catch (IOException e) {
            e.printStackTrace();
        }
        LOGGER.info("EEC read finished. used: {}", System.currentTimeMillis() - start);
    }

    private void eecRead0(String name) {
        LOGGER.info("EEC start to read...");
        long start = System.currentTimeMillis();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve(name + ".xlsx"))) {
            long n = reader.sheets().flatMap(Sheet::rows).count();
            LOGGER.info("Data rows: {}", n);
        } catch (IOException e) {
            e.printStackTrace();
        }
        LOGGER.info("EEC read finished. used: {}", System.currentTimeMillis() - start);
    }

    private void eecWriteShared(String name) throws IOException {
        LOGGER.info("EEC start to write...");
        long start = System.currentTimeMillis();
        new Workbook().addSheet(new ListSheet<LargeSharedData>() {
            int n = 0;
            public List<LargeSharedData> more() {
                LOGGER.info("{} fill success.", n);
                return n++ < loop ? createSharedData() : null;
            }
        }).writeTo(defaultTestPath.resolve(name + ".xlsx"));
        LOGGER.info("EEC write finished. used: {}", System.currentTimeMillis() - start);
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

    private static String[] provinces = {"江苏省", "湖北省", "浙江省", "广东省"};
    private static String[][] cities = {{"南京市", "苏州市", "无锡市", "徐州市"}
        , {"武汉市", "黄冈市", "黄石市", "孝感市", "宜昌市"}
        , {"杭州市", "温州市", "绍兴市", "嘉兴市"}
        , {"广州市", "深圳市", "佛山市"}
    };
    private static String[][][] areas = {{
        {"玄武区", "秦淮区", "鼓楼区", "雨花台区", "栖霞区"}
        , {"虎丘区", "吴中区", "相城区", " 姑苏区", "吴江区"}
        , {"锡山区", "惠山区", "滨湖区", "新吴区", "江阴市"}
        , {"鼓楼区", "云龙区", "贾汪区", "泉山区"}
    }, {
        {"江岸区", "江汉区", "硚口区", "汉阳区", "武昌区", "青山区", "洪山区", "东西湖区"}
        , {"黄州区", "团风县", "红安县"}
        , {"黄石港区", "西塞山区", "下陆区", "铁山区"}
        , {"孝南区", "孝昌县", "大悟县", "云梦县"}
        , {"西陵区", "伍家岗区", "点军区"}
    }, {
        {"上城区", "下城区", "江干区", "拱墅区", "西湖区", "滨江区", "余杭区", "萧山区"}
        , {"鹿城区", "龙湾区", "洞头区"}
        , {"越城区", "柯桥区", "上虞区", "新昌县", "诸暨市", "嵊州市"}
        , {"南湖区", "秀洲区", "嘉善县", "海盐县", "海宁市", "平湖市", "桐乡市"}
    }, {
        {"荔湾区", "白云区", "天河区", "黄埔区", "番禺区", "花都区"}
        , {"罗湖区", "福田区", "南山区", "龙岗区"}
        , {"禅城区", "南海区", "顺德区", "三水区", "高明区"}
    }};
    private List<LargeSharedData> createSharedData() {
        List<LargeSharedData> list = new ArrayList<>();
        int size = i + 1000, p, c;
        for (; i < size; i++) {
            LargeSharedData largeData = new LargeSharedData();
            list.add(largeData);
            largeData.setNv(random.nextInt());
            largeData.setLv(random.nextLong());
            largeData.setDv(random.nextDouble());
            largeData.setAv(new Date(System.currentTimeMillis() - random.nextInt(9999)));
            largeData.setProvince(provinces[p = random.nextInt(provinces.length)]);
            largeData.setCity(cities[p][c = random.nextInt(cities[p].length)]);
            largeData.setArea(areas[p][c][random.nextInt(areas[p][c].length)]);
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