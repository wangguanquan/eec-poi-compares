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

import java.util.ArrayList;
import java.util.List;

import static org.ttzero.compares.BaseExcelTest.getRandomString;
import static org.ttzero.compares.BaseExcelTest.random;

/**
 * @author guanquan.wang at 2020-03-05 16:59
 */
public class Student {
    private int id;
    @ExcelColumn("姓名")
    private String name;
    @ExcelColumn("成绩")
    private int score;

    protected Student(int id, String name, int score) {
        this.id = id;
        this.name = name;
        this.score = score;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getScore() {
        return score;
    }

    public void setScore(int score) {
        this.score = score;
    }

    public static List<Student> randomTestData(int pageNo, int limit) {
        List<Student> list = new ArrayList<>(limit);
        for (int i = pageNo * limit, n = i + limit; i < n; i++) {
            Student e = new Student(i, getRandomString(), random.nextInt(50) + 50);
            list.add(e);
        }
        return list;
    }

    public static List<Student> randomTestData(int n) {
        return randomTestData(0, n);
    }

    public static List<Student> randomTestData() {
        int n = random.nextInt(100) + 1;
        return randomTestData(n);
    }

    @Override
    @ExcelColumn
    public String toString() {
        return "id: " + id + ", name: " + name + ", score: " + score;
    }
}
