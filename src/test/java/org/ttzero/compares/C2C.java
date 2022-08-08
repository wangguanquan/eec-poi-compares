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

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import org.ttzero.excel.annotation.ExcelColumn;

/**
 * @author guanquan.wang at 2020-03-05 11:20
 */
public class C2C {
    @ExcelColumn
    @ExcelIgnore
    @ExcelProperty
    private int ch_id;
    @ExcelColumn
    private int cu_id;
    public C2C() { }
    public C2C(int ch_id, int cu_id) {
        this.ch_id = ch_id;
        this.cu_id = cu_id;
    }

    public int getCh_id() {
        return ch_id;
    }

    public void setCh_id(int ch_id) {
        this.ch_id = ch_id;
    }

    public int getCu_id() {
        return cu_id;
    }

    public void setCu_id(int cu_id) {
        this.cu_id = cu_id;
    }
}
