package com.abnote.planilhas.estilos.estilos;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class CenterStyle {
    private final Workbook workbook;
    private final Sheet sheet;
    private final Map<String, CellStyle> styleCache;

    public CenterStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.styleCache = styleCache;
    }

    public void centralizarTudo(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex,
            boolean isRange) {
        if (isRange) {
            for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null)
                    continue;
                for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
                    Cell cell = row.getCell(colIdx);
                    if (cell == null)
                        continue;
                    CellStyle originalStyle = cell.getCellStyle();
                    CellStyle newStyle = createCombinedStyle(originalStyle, false, true);
                    cell.setCellStyle(newStyle);
                }
            }
        } else {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell != null) {
                        CellStyle originalStyle = cell.getCellStyle();
                        CellStyle newStyle = createCombinedStyle(originalStyle, false, true);
                        cell.setCellStyle(newStyle);
                    }
                }
            }
        }
    }

    public void centralizarERedimensionarTudo() {
        // Centralização
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell != null) {
                    CellStyle originalStyle = cell.getCellStyle();
                    CellStyle newStyle = workbook.createCellStyle();
                    newStyle.cloneStyleFrom(originalStyle);
                    newStyle.setAlignment(HorizontalAlignment.CENTER);
                    newStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    cell.setCellStyle(newStyle);
                }
            }
        }

        // Redimensionamento das colunas
        redimensionarColunas();
    }

    private CellStyle createCombinedStyle(CellStyle originalStyle, boolean addBold, boolean addCentered) {
        CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(originalStyle);

        if (addBold) {
            // Implementação se necessário
        }

        if (addCentered) {
            newStyle.setAlignment(HorizontalAlignment.CENTER);
            newStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }

        return newStyle;
    }

    public void redimensionarColunas() {
        int maxColumns = getMaxNumberOfColumns();
        for (int i = 0; i < maxColumns; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private int getMaxNumberOfColumns() {
        int maxColumns = 0;
        for (Row row : sheet) {
            if (row.getLastCellNum() > maxColumns) {
                maxColumns = row.getLastCellNum();
            }
        }
        return maxColumns;
    }
}
