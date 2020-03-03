package org.ttzero.compares;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Jiaju Zhuang
 */
public class LargeDataListener extends AnalysisEventListener<LargeData> {
    private static final Logger LOGGER = LoggerFactory.getLogger(LargeDataListener.class);
    private int count = 0;

    @Override
    public void invoke(LargeData data, AnalysisContext context) {
        count++;
        if (count % 100000 == 0) {
            LOGGER.info("Already read:{}", count);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        LOGGER.info("Large row count:{}", count);
    }
}
