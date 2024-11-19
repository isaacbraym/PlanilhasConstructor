package com.abnote.planilhas.estilos.estilos;

import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.utils.PosicaoConverter;

public class BorderStyleHelper {
    private final Workbook workbook;
    private final Sheet sheet;
    private final Map<String, CellStyle> styleCache;

    public BorderStyleHelper(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.styleCache = styleCache;
    }

    public void todasAsBordasEmTudo(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex,
            boolean isRange) {
        if (isRange) {
            for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null)
                    continue;
                for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
                    Cell cell = row.getCell(colIdx);
                    if (cell == null || cell.getCellTypeEnum() == CellType.BLANK) {
                        continue; // Pular células vazias
                    }
                    // Obter o estilo atual da célula
                    CellStyle originalStyle = cell.getCellStyle();

                    // Verificar se a célula já possui alguma borda espessa
                    boolean hasThickBorder = originalStyle.getBorderTopEnum() == BorderStyle.THICK
                            || originalStyle.getBorderBottomEnum() == BorderStyle.THICK
                            || originalStyle.getBorderLeftEnum() == BorderStyle.THICK
                            || originalStyle.getBorderRightEnum() == BorderStyle.THICK;

                    if (hasThickBorder) {
                        continue; // Ignorar células com bordas espessas
                    }

                    // Gerar uma chave única para o cache baseado no estilo original
                    String cacheKey = "borders_" + originalStyle.hashCode();
                    CellStyle borderedStyle = styleCache.get(cacheKey);
                    if (borderedStyle == null) {
                        // Clonar o estilo original
                        borderedStyle = workbook.createCellStyle();
                        borderedStyle.cloneStyleFrom(originalStyle);
                        // Aplicar bordas finas
                        borderedStyle.setBorderTop(BorderStyle.THIN);
                        borderedStyle.setBorderBottom(BorderStyle.THIN);
                        borderedStyle.setBorderLeft(BorderStyle.THIN);
                        borderedStyle.setBorderRight(BorderStyle.THIN);
                        // Armazenar no cache
                        styleCache.put(cacheKey, borderedStyle);
                    }
                    // Aplicar o estilo com bordas à célula
                    cell.setCellStyle(borderedStyle);
                }
            }
        } else {
            // Comportamento existente para aplicar bordas a toda a planilha
            for (Row row : sheet) {
                if (row == null)
                    continue;
                for (Cell cell : row) {
                    if (cell == null || cell.getCellTypeEnum() == CellType.BLANK) {
                        continue; // Pular células vazias
                    }
                    // Obter o estilo atual da célula
                    CellStyle originalStyle = cell.getCellStyle();

                    // Verificar se a célula já possui alguma borda espessa
                    boolean hasThickBorder = originalStyle.getBorderTopEnum() == BorderStyle.THICK
                            || originalStyle.getBorderBottomEnum() == BorderStyle.THICK
                            || originalStyle.getBorderLeftEnum() == BorderStyle.THICK
                            || originalStyle.getBorderRightEnum() == BorderStyle.THICK;

                    if (hasThickBorder) {
                        continue; // Ignorar células com bordas espessas
                    }

                    // Gerar uma chave única para o cache baseado no estilo original
                    String cacheKey = "borders_" + originalStyle.hashCode();
                    CellStyle borderedStyle = styleCache.get(cacheKey);
                    if (borderedStyle == null) {
                        // Clonar o estilo original
                        borderedStyle = workbook.createCellStyle();
                        borderedStyle.cloneStyleFrom(originalStyle);
                        // Aplicar bordas finas
                        borderedStyle.setBorderTop(BorderStyle.THIN);
                        borderedStyle.setBorderBottom(BorderStyle.THIN);
                        borderedStyle.setBorderLeft(BorderStyle.THIN);
                        borderedStyle.setBorderRight(BorderStyle.THIN);
                        // Armazenar no cache
                        styleCache.put(cacheKey, borderedStyle);
                    }
                    // Aplicar o estilo com bordas à célula
                    cell.setCellStyle(borderedStyle);
                }
            }
        }
    }

    public void aplicarBordasNaCelula(String posicao) {
        int[] posicaoIndices = PosicaoConverter.converterPosicao(posicao);
        int coluna = posicaoIndices[0];
        int linha = posicaoIndices[1];
        Row row = sheet.getRow(linha);
        if (row == null) {
            row = sheet.createRow(linha);
        }
        Cell cell = row.getCell(coluna);
        if (cell == null) {
            cell = row.createCell(coluna);
        }
        // Preservar estilos existentes
        CellStyle originalStyle = cell.getCellStyle();
        CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(originalStyle);
        newStyle.setBorderTop(BorderStyle.THIN);
        newStyle.setBorderBottom(BorderStyle.THIN);
        newStyle.setBorderLeft(BorderStyle.THIN);
        newStyle.setBorderRight(BorderStyle.THIN);
        cell.setCellStyle(newStyle);
    }

    public void aplicarTodasAsBordasDeAte(String posicaoInicial, String posicaoFinal) {
        int[] inicio = PosicaoConverter.converterPosicao(posicaoInicial);
        int[] fim = PosicaoConverter.converterPosicao(posicaoFinal);

        for (int rowIdx = inicio[1]; rowIdx <= fim[1]; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) {
                row = sheet.createRow(rowIdx);
            }
            for (int colIdx = inicio[0]; colIdx <= fim[0]; colIdx++) {
                Cell cell = row.getCell(colIdx);
                if (cell == null) {
                    cell = row.createCell(colIdx);
                }
                CellStyle borderStyle = workbook.createCellStyle();
                borderStyle.setBorderTop(BorderStyle.THIN);
                borderStyle.setBorderBottom(BorderStyle.THIN);
                borderStyle.setBorderLeft(BorderStyle.THIN);
                borderStyle.setBorderRight(BorderStyle.THIN);
                cell.setCellStyle(borderStyle);
            }
        }
    }

    public void bordasEspessas(String posicaoInicial, String posicaoFinal) {
        int[] inicio = PosicaoConverter.converterPosicao(posicaoInicial);
        int[] fim = PosicaoConverter.converterPosicao(posicaoFinal);

        for (int rowIdx = inicio[1]; rowIdx <= fim[1]; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) {
                row = sheet.createRow(rowIdx);
            }
            for (int colIdx = inicio[0]; colIdx <= fim[0]; colIdx++) {
                Cell cell = row.getCell(colIdx);
                if (cell == null) {
                    cell = row.createCell(colIdx);
                }

                CellStyle originalStyle = cell.getCellStyle();
                String cacheKey = "thickBorders_" + originalStyle.hashCode() + "_r" + rowIdx + "_c" + colIdx;
                CellStyle thickBorderStyle = styleCache.get(cacheKey);
                if (thickBorderStyle == null) {
                    thickBorderStyle = workbook.createCellStyle();
                    thickBorderStyle.cloneStyleFrom(originalStyle);

                    // Aplicar bordas espessas nas bordas externas
                    if (rowIdx == inicio[1]) { // Primeira linha do intervalo
                        thickBorderStyle.setBorderTop(BorderStyle.THICK);
                    }
                    if (rowIdx == fim[1]) { // Última linha do intervalo
                        thickBorderStyle.setBorderBottom(BorderStyle.THICK);
                    }
                    if (colIdx == inicio[0]) { // Primeira coluna do intervalo
                        thickBorderStyle.setBorderLeft(BorderStyle.THICK);
                    }
                    if (colIdx == fim[0]) { // Última coluna do intervalo
                        thickBorderStyle.setBorderRight(BorderStyle.THICK);
                    }

                    // Armazenar no cache para reutilização
                    styleCache.put(cacheKey, thickBorderStyle);
                }

                // Aplicar o estilo com bordas espessas à célula
                cell.setCellStyle(thickBorderStyle);
            }
        }
    }

    public void bordasEspessasComBordasInternas(String posicaoInicial, String posicaoFinal) {
        int[] inicio = PosicaoConverter.converterPosicao(posicaoInicial);
        int[] fim = PosicaoConverter.converterPosicao(posicaoFinal);

        for (int rowIdx = inicio[1]; rowIdx <= fim[1]; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) {
                row = sheet.createRow(rowIdx);
            }
            for (int colIdx = inicio[0]; colIdx <= fim[0]; colIdx++) {
                Cell cell = row.getCell(colIdx);
                if (cell == null) {
                    cell = row.createCell(colIdx);
                }

                CellStyle originalStyle = cell.getCellStyle();
                StringBuilder cacheKeyBuilder = new StringBuilder("thickInternalBorders_");
                cacheKeyBuilder.append(originalStyle.hashCode()).append("_r").append(rowIdx).append("_c")
                        .append(colIdx);
                String cacheKey = cacheKeyBuilder.toString();

                CellStyle newStyle = styleCache.get(cacheKey);
                if (newStyle == null) {
                    newStyle = workbook.createCellStyle();
                    newStyle.cloneStyleFrom(originalStyle);

                    // Aplicar bordas finas em todas as direções
                    newStyle.setBorderTop(BorderStyle.THIN);
                    newStyle.setBorderBottom(BorderStyle.THIN);
                    newStyle.setBorderLeft(BorderStyle.THIN);
                    newStyle.setBorderRight(BorderStyle.THIN);

                    // Aplicar bordas espessas nas bordas externas
                    if (rowIdx == inicio[1]) { // Primeira linha do intervalo
                        newStyle.setBorderTop(BorderStyle.THICK);
                    }
                    if (rowIdx == fim[1]) { // Última linha do intervalo
                        newStyle.setBorderBottom(BorderStyle.THICK);
                    }
                    if (colIdx == inicio[0]) { // Primeira coluna do intervalo
                        newStyle.setBorderLeft(BorderStyle.THICK);
                    }
                    if (colIdx == fim[0]) { // Última coluna do intervalo
                        newStyle.setBorderRight(BorderStyle.THICK);
                    }

                    // Armazenar no cache para reutilização
                    styleCache.put(cacheKey, newStyle);
                }

                // Aplicar o novo estilo à célula
                cell.setCellStyle(newStyle);
            }
        }
    }
}
