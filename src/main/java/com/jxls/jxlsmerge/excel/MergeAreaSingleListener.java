package com.jxls.jxlsmerge.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;

import java.util.List;
import java.util.Objects;

@Slf4j
public class MergeAreaSingleListener implements AreaListener {

    private final CellRef commandCell;

    private final Sheet sheet;

    private final String item;

    private CellRef lastRowCellRef;

    private XSSFWorkbook workbook;

    public MergeAreaSingleListener(Transformer transformer, CellRef cellRef, String item) {
        this.commandCell = cellRef;
        this.workbook = ((PoiTransformer) transformer).getXSSFWorkbook();
        this.sheet = workbook.getSheet(cellRef.getSheetName());
        this.item = item;
    }


    @Override
    public void afterApplyAtCell(CellRef cellRef, Context context) {

        List<String> object = (List<String>) context.getVar(item);
        int totalSize = object.size();
        int curSize = (int) context.getVar("index");

        if(curSize + 1 == totalSize || !Objects.equals(object.get(curSize), object.get(curSize + 1))) {
            merge(cellRef);
        }else{
            return;
        }

    }

    private void merge(CellRef cellRef) {
        // 병합이 시작될 지점
        int from = this.lastRowCellRef == null ? commandCell.getRow() : this.lastRowCellRef.getRow() + 1;
        int lastRow = cellRef.getRow();

        this.setLastRowCellRef(cellRef);

        log.debug("this :{}, merged start row : {} | end row : {} | col :{} ", this.toString(), from, lastRow, cellRef.getCol());

        if(from != lastRow) {
            // 병합 cell 생성
            CellRangeAddress region = new CellRangeAddress(from, lastRow, cellRef.getCol(), cellRef.getCol());

            // sheet 에 병합된 셀 추가
            sheet.addMergedRegion(region);
        }

        // 스타일 적용
        applyStyle(sheet.getRow(from).getCell(cellRef.getCol()));
    }

    private void setLastRowCellRef(CellRef cellRef) {
        if (this.lastRowCellRef == null || this.lastRowCellRef.getRow() < cellRef.getRow()) {
            this.lastRowCellRef = cellRef;
        }
    }

    private void applyStyle(Cell cell) {
        CellStyle cellStyle = null;
        Font font = workbook.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());

        // 배경, 폰트 지정
        if(cell.toString().equals("불합격")) {
            cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);
            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }else if(cell.toString().equals("합격")) {
            cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);
            cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }else{
            cellStyle = cell.getCellStyle();
        }

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(cellStyle);
    }


    private boolean existMerged(CellRef cell) {
        return sheet.getMergedRegions().stream()
                .anyMatch(address -> address.isInRange(cell.getRow(), cell.getCol()));
    }

    @Override
    public void beforeApplyAtCell(CellRef cellRef, Context context) {
    }

    @Override
    public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
    }

    @Override
    public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
    }

}


