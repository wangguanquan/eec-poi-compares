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
import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
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
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Styles;
import org.ttzero.excel.reader.Dimension;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.reader.Sheet;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.AccessibleObject;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.UUID;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static org.ttzero.compares.LargeExcelTest.defaultTestPath;
import static org.ttzero.excel.util.StringUtil.EMPTY;
import static org.ttzero.excel.util.StringUtil.isNotEmpty;

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
//        new Workbook("test8")
//            .addSheet(new ListSheet<>("帐单表", checks()))
//            .addSheet(new ListSheet<>("客户表", customers()).hidden())
//            .addSheet(new ListSheet<>("用户客户关系表", c2CS()))
//            .writeTo(defaultTestPath);
        String fileName = "abc";
        String path = "/root/" + System.currentTimeMillis() + "-" + UUID.randomUUID().toString();
        int n = fileName.lastIndexOf('.');
        File dirFile = new File(path + (n > 0 ? fileName.substring(n) : ""));
        System.out.println(dirFile);
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
                .setProcessor(n -> (int)n < 60 ? "不及格" : n)
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
            reader.sheets()
                .peek(sheet -> System.out.println("----------" + sheet.getName() + "-----------"))
                .flatMap(Sheet::rows)
                .forEach(System.out::println);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test public void test11() {
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test8.xlsx"))) {
            Sheet firstSheet = reader.sheet(0);

            Dimension dimension = firstSheet.getDimension();
            // lastRow - firstRow = 数据行的行数，不包含header
            if (dimension.lastRow - dimension.firstRow > 1000) {
                // 如果数据量超过1千则选择流式处理，forEach里也可以收集一定量的实体再批量处理
                firstSheet.dataRows().map(row -> row.too(Check.class)).forEach(check -> {
                    // TODO 业务处理
                });
            } else {
                // 数据量小于1千则直接转为集合处理
                List<Check> checks = firstSheet.dataRows().map(row -> row.to(Check.class)).collect(Collectors.toList());
                // TODO 业务处理
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static class EasyExcelSupportListSheet<T> extends ListSheet<T> {
        /**
         * 过滤不需要导出的字段
         *
         * @param ao {@link T}对象定义的所有{@link java.lang.reflect.Field}和 {@link java.lang.reflect.Method}
         * @return true: 忽略该字段（该字段不导出）
         */
        @Override
        protected boolean ignoreColumn(AccessibleObject ao) {
            // Easyexcel使用注解com.alibaba.excel.annotation.ExcelIgnore来标记忽略字段
            return ao.getAnnotation(ExcelIgnore.class) != null;
        }

        /**
         * 创建列
         *
         * @param ao {@link T}对象定义的所有{@link java.lang.reflect.Field}和 {@link java.lang.reflect.Method}
         * @return 列定义，可以返回null，表示忽略该字段
         */
        @Override
        protected ListSheet.EntryColumn createColumn(AccessibleObject ao) {
            // 1. 过滤掉需要忽略的字段
            if (ignoreColumn(ao)) return null;

            ao.setAccessible(true);
            // 2. Easyexcel使用com.alibaba.excel.annotation.ExcelProperty标记，这里可以替换为任意的定义的注解，当然需要包含一些列头的信息
            ExcelProperty ec = ao.getAnnotation(ExcelProperty.class);
            if (ec != null) {
                ListSheet.EntryColumn column = new ListSheet.EntryColumn(ec.value()[0], EMPTY);
            /*
             Easyexcel格式化有3种方式
             一是放在{@code ExcelProperty#format} 注解
             二是使用DateTimeFormat和NumberFormat注解
             所以这里要兼容这3种方式

             EEC统一使用NumFmt设计
             */
                DateTimeFormat dateTimeFormat = ao.getAnnotation(DateTimeFormat.class);
                NumberFormat numberFormat = ao.getAnnotation(NumberFormat.class);
                // 3. 格式化（支持日期和数字）
                if (isNotEmpty(ec.format())) {
                    column.setNumFmt(ec.format());
                } else if (dateTimeFormat != null) {
                    column.setNumFmt(dateTimeFormat.value());
                } else if (numberFormat != null) {
                    column.setNumFmt(numberFormat.value());
                }
                // 4. 列位置
                if (ec.index() > -1) {
                    column.setColIndex(ec.index());
                }
                // 5. 列宽
                ColumnWidth columnWidth = ao.getAnnotation(ColumnWidth.class);
                if (columnWidth != null && columnWidth.value() > 0) {
                    column.width = columnWidth.value();
                }
                return column;

                // TODO 其它属性
            }
            return null;
        }
    }

    @Test
    public void testReadText() {
        try (Stream<String> line = Files.lines(Paths.get("./1.txt"))) {
            line.forEach(System.out::println);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    private RedLineDetectService redLineDetectService;
    private GoodsService goodsService;

    @Test
    public void testReadExcel() {
        // 假定已将文件写到某台服务器
        try (ExcelReader reader = ExcelReader.read(Paths.get("./goods.xlsx"))) {
            reader.sheet(0).dataRows()
                // 将行数据转为Goods对象，可以使用{@code row#to} 或者 {@code row#too} 两种方式，后者内存共享，如果要调用collect方法则必须使用第一种方式，否则仅有最后一行数据。
                .map(row -> row.too(Goods.class))

                // 本地的一些检查，检查一些必填项
//            .filter(this::validateGoods)

                // 合规检查，检查文本或者图片是否违规
                .filter(g -> redLineDetectService.checkText(g.getGoodsName()) && redLineDetectService.checkImage(g.getImage()))

                // 调用商品服务上架
                .forEach(goodsService::publish);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }


    @Test
    public void testExcelForeach() {
        try (ExcelReader reader = ExcelReader.read(Paths.get("./goods.xlsx"))) {
            List<Goods> batch = new ArrayList<>(100);
            for (Iterator<Row> ite = reader.sheet(0).dataIterator(); ite.hasNext(); ) {
                // 行数据转对象
                batch.add(ite.next().to(Goods.class));
                // 满100条批量上架
                if (batch.size() >= 100) {
                    goodsService.batchPublish(batch);
                    batch.clear();
                }
            }
            // 上架剩余商品
            if (!batch.isEmpty()) {
                goodsService.batchPublish(batch);
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    @Test
    public void testReset() {
        try (ExcelReader reader = ExcelReader.read(Paths.get("./large-goods.xlsx"))) {
            Sheet sheet = reader.sheet(0);
//        Dimension dimension = sheet.getDimension();
//        // 这里没有+1，因为表头占一行
//        int maxRow = dimension.lastRow - dimension.firstRow;
//        long count = sheet.dataRows().map(row -> row.getString("商品编码")).distinct().count();
//        // 如果去重后的结果小于原始结果说明有重复
//        if (count < maxRow) {
//            throw new IllegalArgumentException("包含重复商品编码");
//        }
//
//        // 重置
//        sheet.reset();
//
//        // 和上面的例子一样处理逻辑即可
//        sheet.dataRows().map(row -> row.to(Goods.class)).forEach(goodsService::publish);

            sheet.dataRows().collect(Collectors.toMap(row -> row.getString(0), row -> row.getString(1), (a, b) -> a));
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }


    @Test
    public void testAutoSize() throws IOException {
        List<String> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            list.add("联想笔记本电脑拯救者黑");
        }
        new Workbook().setAutoSize(true).addSheet(new ListSheet<String>(list) {
            @Override
            protected org.ttzero.excel.entity.Column[] getHeaderColumns() {
                return new org.ttzero.excel.entity.Column[]{new EntryColumn() {
                    @Override
                    public int getCellStyle(Class<?> clazz) {
                        return styles.addFont(new Font("宋体", 20));
                    }
                }.setClazz(String.class)};
            }
        }.ignoreHeader().cancelOddStyle().setSheetWriter(new MyXMLWorksheetWriter())).writeTo(defaultTestPath);
    }

    public static class MyXMLWorksheetWriter extends XMLWorksheetWriter {
        @Override
        protected double stringWidth(String s, int xf) {
            if (StringUtil.isEmpty(s)) return 0.0D;
            int n = Math.min(s.length(), cacheChar.length);
            double w = 0.0D;
            s.getChars(0, n, cacheChar, 0);

            // 通过xf来获取当前字符串对应的字体
            Styles styles = sheet.getWorkbook().getStyles();
            int style = styles.getStyleByIndex(xf);
            Font font = styles.getFont(style);

            double doubleByte = 0.0D; // 双字节大小
            double singleByte = 0.0D; // 单字节大小
        /*
        在这里进行微调，通过字体，大小来调整单子节和双字节的大小
         */
            if ("宋体".equals(font.getName())) {
                if (font.getSize() < 12) {
                    singleByte = 1.0D;
                    doubleByte = 1.86D;
                } else if (font.getSize() < 15) {
                    singleByte = 1.26D; // 请微调
                    doubleByte = 2.16D;
                } else if (font.getSize() < 18) {
                    singleByte = 1.26D; // 请微调
                    doubleByte = 2.26D;
                } else if (font.getSize() < 21) {
                    singleByte = 1.73D;
                    doubleByte = 3.46D;
                } else {
                    singleByte = 2.16D;
                    doubleByte = 3.26D;
                }
            }
            for (int i = 0; i < n; w += cacheChar[i++] > 0x4E00 ? doubleByte : singleByte) ;
            return w;
        }
    }























//    private boolean validateGoods(Goods goods) {
//        return StringUtils.isNotEmpty(goods.getGoodsName())
//            && StringUtils.isNotEmpty(goods.getImage());
//    }





    public static class Goods {
        public String getGoodsName() { return null; }
        public String getImage() { return null; }
        public String getGoodsNo() { return null; }
    }

    public abstract static class RedLineDetectService {
        public abstract boolean checkText(String txt);
        public abstract boolean checkImage(String url);
    }

    public abstract static class GoodsService {
        public abstract void publish(Goods goods);
        public abstract void batchPublish(List<Goods> goods);
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
