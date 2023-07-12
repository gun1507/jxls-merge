package com.jxls.jxlsmerge.excel;

import com.jxls.jxlsmerge.domain.Pass;
import com.jxls.jxlsmerge.domain.School;
import lombok.extern.slf4j.Slf4j;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.List;

@Component
@Slf4j
public class ExcelUtil {

    public ExcelUtil() {
        XlsCommentAreaBuilder.addCommandMapping(EachMergeCommand.COMMAND_NAME, EachMergeCommand.class);
        XlsCommentAreaBuilder.addCommandMapping(EachMergeSingleCommand.COMMAND_NAME, EachMergeSingleCommand.class);
    }

    public void writeExcelFile(String outputPath, List<School> schoolList, List<Pass> passList) throws Exception {

        // 템플릿 저장장소
        ClassLoader classLoader = getClass().getClassLoader();
        InputStream is = classLoader.getResourceAsStream("excel/template.xlsx");
        OutputStream os = null;

        try {

            Context context = new Context();
            os = new FileOutputStream(outputPath);
            context.putVar("schools", schoolList);
            context.putVar("passes", passList);
            JxlsHelper.getInstance().processTemplate(is, os, context);
        } catch (IOException e) {
            throw new ExcelGenerateException("엑셀 파일 생성 실패 하였습니다.", e);
        } finally {
            os.close();
        }
    }
}
