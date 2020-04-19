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

import org.junit.Test;

import static org.ttzero.compares.LargeExcelTest.easySharedRead;
import static org.ttzero.compares.LargeExcelTest.eecSharedRead;

/**
 * @author guanquan.wang at 2020-03-23 11:57
 */
public class XlsTest {

    @Test public void testEec1w() {
        eecSharedRead("eec shared 1w.xls");
    }

    @Test public void testEec5w() {
        eecSharedRead("eec shared 5w.xls");
    }

    @Test public void testEec10w() {
        eecSharedRead("eec shared 10w.xls");
    }

    @Test public void testEec32w() {
        eecSharedRead("eec shared 32w.xls");
    }

    @Test public void testEas1w() {
        easySharedRead("eec shared 1w.xls");
    }

    @Test public void testEas5w() {
        easySharedRead("eec shared 5w.xls");
    }

    @Test public void testEas10w() {
        easySharedRead("eec shared 10w.xls");
    }

    @Test public void testEas32w() {
        easySharedRead("eec shared 32w.xls");
    }
}
