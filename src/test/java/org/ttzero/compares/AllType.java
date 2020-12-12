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

import org.ttzero.excel.annotation.ExcelColumn;

import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static org.ttzero.compares.BaseExcelTest.charArray;
import static org.ttzero.compares.BaseExcelTest.getRandomString;
import static org.ttzero.compares.BaseExcelTest.random;

/**
 * @author guanquan.wang at 2020-03-05 15:41
 */
public class AllType {
    @ExcelColumn
    private boolean bv;
    @ExcelColumn
    private char cv;
    @ExcelColumn
    private short sv;
    @ExcelColumn
    private int nv;
    @ExcelColumn
    private long lv;
    @ExcelColumn
    private float fv;
    @ExcelColumn
    private double dv;
    @ExcelColumn
    private String s;
    @ExcelColumn
    private BigDecimal mv;
    @ExcelColumn
    private Date av;
    @ExcelColumn
    private Timestamp iv;
    @ExcelColumn
    private Time tv;
    @ExcelColumn
    private LocalDate ldv;
    @ExcelColumn
    private LocalDateTime ldtv;
    @ExcelColumn
    private LocalTime ltv;

    public static List<AllType> randomTestData(int size) {
        List<AllType> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            AllType o = new AllType();
            o.bv = random.nextInt(10) == 5;
            o.cv = charArray[random.nextInt(charArray.length)];
            o.sv = (short) (random.nextInt() & 0xFFFF);
            o.nv = random.nextInt();
            o.lv = random.nextLong();
            o.fv = random.nextFloat();
            o.dv = random.nextDouble();
            o.s = getRandomString();
            o.mv = BigDecimal.valueOf(random.nextDouble());
            o.av = new Date();
            o.iv = new Timestamp(System.currentTimeMillis() - random.nextInt(9999));
            o.tv = new Time(random.nextLong());
            o.ldv = LocalDate.now();
            o.ldtv = LocalDateTime.now();
            o.ltv = LocalTime.now();
            list.add(o);
        }
        return list;
    }

    public static List<AllType> randomTestData() {
        int size = random.nextInt(100) + 1;
        return randomTestData(size);
    }

    public boolean isBv() {
        return bv;
    }

    public char getCv() {
        return cv;
    }

    public short getSv() {
        return sv;
    }

    public int getNv() {
        return nv;
    }

    public long getLv() {
        return lv;
    }

    public float getFv() {
        return fv;
    }

    public double getDv() {
        return dv;
    }

    public String getS() {
        return s;
    }

    public BigDecimal getMv() {
        return mv;
    }

    public Date getAv() {
        return av;
    }

    public Timestamp getIv() {
        return iv;
    }

    public Time getTv() {
        return tv;
    }

    public LocalDate getLdv() {
        return ldv;
    }

    public LocalDateTime getLdtv() {
        return ldtv;
    }

    public LocalTime getLtv() {
        return ltv;
    }
}
