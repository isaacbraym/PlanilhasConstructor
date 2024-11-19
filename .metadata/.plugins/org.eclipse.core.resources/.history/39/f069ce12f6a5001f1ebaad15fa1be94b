package com.abnote.planilhas.estilos.estilos;

import java.util.Map;
import java.util.Objects;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class BoldStyle {

    private final Workbook workbook;
    private final Sheet sheet;
    private final Map<String, CellStyle> styleCache;

    public BoldStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.styleCache = styleCache;
    }

    public void aplicarBold(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
                            int endRowIndex, int endColumnIndex, boolean isRange) {
        if (isRange) {
            aplicarBoldEmIntervalo(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        } else if (rowIndex != -1) {
            if (columnIndex == -1) {
                aplicarBoldEmLinha(rowIndex);
            } else {
                aplicarBoldEmCelulaEspecifica(rowIndex, columnIndex);
            }
        }
    }

    private void aplicarBoldEmIntervalo(int startRowIndex, int startColumnIndex,
                                        int endRowIndex, int endColumnIndex) {
        for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) continue;

            for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
                Cell cell = row.getCell(colIdx);
                if (cell == null) continue;
                aplicarBoldNaCelula(cell);
            }
        }
    }

    private void aplicarBoldEmLinha(int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            for (Cell cell : row) {
                if (cell != null) {
                    aplicarBoldNaCelula(cell);
                }
            }
        }
    }

    private void aplicarBoldEmCelulaEspecifica(int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                aplicarBoldNaCelula(cell);
            }
        }
    }

    private void aplicarBoldNaCelula(Cell cell) {
        CellStyle currentStyle = cell.getCellStyle();
        Font currentFont = workbook.getFontAt(currentStyle.getFontIndex());

        Font boldFont = encontrarOuCriarFonteBold(currentFont);

        String styleKey = "bold_" + currentStyle.hashCode();
        CellStyle newStyle = styleCache.get(styleKey);
        if (newStyle == null) {
            newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(currentStyle);
            newStyle.setFont(boldFont);
            styleCache.put(styleKey, newStyle);
        }

        cell.setCellStyle(newStyle);
    }

    private Font encontrarOuCriarFonteBold(Font currentFont) {
        Font boldFont = buscarFonteBoldExistente(currentFont);
        if (boldFont != null) {
            return boldFont;
        } else {
            return criarFonteBold(currentFont);
        }
    }

    private Font buscarFonteBoldExistente(Font currentFont) {
        for (short i = 0; i < workbook.getNumberOfFonts(); i++) {
            Font font = workbook.getFontAt(i);
            if (font.getBold() && fontesIguaisExcetoBold(font, currentFont)) {
                return font;
            }
        }
        return null;
    }

    private Font criarFonteBold(Font currentFont) {
        Font newFont = workbook.createFont();
        copiarAtributosDaFonte(currentFont, newFont);
        newFont.setBold(true);
        return newFont;
    }

    private boolean fontesIguaisExcetoBold(Font font1, Font font2) {
        return font1.getFontName().equals(font2.getFontName())
            && font1.getFontHeightInPoints() == font2.getFontHeightInPoints()
            && font1.getItalic() == font2.getItalic()
            && font1.getStrikeout() == font2.getStrikeout()
            && font1.getTypeOffset() == font2.getTypeOffset()
            && font1.getUnderline() == font2.getUnderline()
            && font1.getCharSet() == font2.getCharSet()
            && fontesComAMesmaCor(font1, font2);
    }

    private void copiarAtributosDaFonte(Font sourceFont, Font targetFont) {
        targetFont.setFontName(sourceFont.getFontName());
        targetFont.setFontHeightInPoints(sourceFont.getFontHeightInPoints());
        targetFont.setItalic(sourceFont.getItalic());
        targetFont.setStrikeout(sourceFont.getStrikeout());
        targetFont.setTypeOffset(sourceFont.getTypeOffset());
        targetFont.setUnderline(sourceFont.getUnderline());
        targetFont.setCharSet(sourceFont.getCharSet());
        targetFont.setColor(sourceFont.getColor());

        if (sourceFont instanceof XSSFFont && targetFont instanceof XSSFFont) {
            XSSFColor color = ((XSSFFont) sourceFont).getXSSFColor();
            ((XSSFFont) targetFont).setColor(color);
        }
    }

    private boolean fontesComAMesmaCor(Font font1, Font font2) {
        if (font1 instanceof XSSFFont && font2 instanceof XSSFFont) {
            XSSFColor color1 = ((XSSFFont) font1).getXSSFColor();
            XSSFColor color2 = ((XSSFFont) font2).getXSSFColor();
            return Objects.equals(color1, color2);
        } else {
            return font1.getColor() == font2.getColor();
        }
    }
}
