package com.jxls.jxlsmerge.excel;

import org.jxls.area.Area;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;

import java.util.List;
import java.util.stream.Collectors;

public class EachMergeCommand extends EachCommand {

    public static final String COMMAND_NAME = "each-merge";

    @Override
    public Size applyAt(CellRef cellRef, Context context) {


        // each-merge 셀 의 하위 셀 영역 목록을 가져온다.
        List<Area> childAreas = this.getAreaList().stream()
                .flatMap(area -> area.getCommandDataList().stream())
                .flatMap(commandData -> commandData.getCommand().getAreaList().stream())
                .collect(Collectors.toList());

        // 셀 병합을 수행하는 MergeAreaListener instance 생성
        MergeAreaListener listener = new MergeAreaListener(this.getTransformer(), cellRef);

        // each-merge comment 가 작성된 cell area 에 MergeAreaListener 추가
        this.getAreaList().get(0).addAreaListener(listener);

        //  하위 영역에 MergeAreaListener 추가
        childAreas.forEach(childArea -> {
            childArea.addAreaListener(listener);
        });

        // each 커맨드 수행
        return super.applyAt(cellRef, context);
    }
}


