package com.abnote.planilhas.calculos;

import org.apache.poi.ss.usermodel.*;
import com.abnote.planilhas.utils.PosicaoConverter;

public class Calculos {

    /**
     * Soma os valores numéricos de uma coluna específica e insere a soma sem texto.
     * A célula da soma mantém a mesma formatação das células somadas.
     *
     * @param sheet          A folha da planilha onde a soma será realizada.
     * @param posicaoInicial A posição inicial da coluna a ser somada (ex: "J3").
     */
    public static void somarColuna(Sheet sheet, String posicaoInicial) {
        int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
        int coluna = posicao[0];
        int linhaInicial = posicao[1];

        double soma = 0.0;
        CellStyle estiloSoma = null;

        Workbook workbook = sheet.getWorkbook();
        int ultimaLinhaDados = -1;

        for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            Cell cell = row.getCell(coluna);
            if (cell != null && cell.getCellTypeEnum() == CellType.NUMERIC) {
                soma += cell.getNumericCellValue();

                if (estiloSoma == null) {
                    estiloSoma = cell.getCellStyle();
                }

                ultimaLinhaDados = i;
            }
        }

        int linhaSoma = ultimaLinhaDados + 1;
        Row rowSoma = sheet.getRow(linhaSoma);
        if (rowSoma == null) {
            rowSoma = sheet.createRow(linhaSoma);
        }

        Cell cellSoma = rowSoma.getCell(coluna);
        if (cellSoma == null) {
            cellSoma = rowSoma.createCell(coluna);
        }
        cellSoma.setCellValue(soma);

        if (estiloSoma != null) {
            CellStyle somaStyle = workbook.createCellStyle();
            somaStyle.cloneStyleFrom(estiloSoma);
            cellSoma.setCellStyle(somaStyle);
        } else {
            CellStyle defaultStyle = workbook.createCellStyle();
            defaultStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
            cellSoma.setCellStyle(defaultStyle);
        }
    }

    /**
     * Soma os valores numéricos de uma coluna específica e insere a soma com um
     * texto descritivo. A célula da soma mantém a mesma formatação das células
     * somadas.
     *
     * @param sheet          A folha da planilha onde a soma será realizada.
     * @param posicaoInicial A posição inicial da coluna a ser somada (ex: "J3").
     * @param texto          O texto descritivo que será inserido ao lado da soma.
     */
    public static void somarColunaComTexto(Sheet sheet, String posicaoInicial, String texto) {
        int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
        int coluna = posicao[0];
        int linhaInicial = posicao[1];

        double soma = 0.0;
        int ultimaLinha = linhaInicial;
        CellStyle estiloSoma = null;

        Workbook workbook = sheet.getWorkbook();

        for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            Cell cell = row.getCell(coluna);
            if (cell != null && cell.getCellTypeEnum() == CellType.NUMERIC) {
                soma += cell.getNumericCellValue();

                if (estiloSoma == null) {
                    estiloSoma = cell.getCellStyle();
                }
            }
            ultimaLinha = i;
        }

        Row linhaSoma = sheet.createRow(ultimaLinha + 1);

        Cell cellTexto = linhaSoma.createCell(coluna - 1);
        cellTexto.setCellValue(texto);

        Cell cellSoma = linhaSoma.createCell(coluna);
        cellSoma.setCellValue(soma);

        if (estiloSoma != null) {
            CellStyle somaStyle = workbook.createCellStyle();
            somaStyle.cloneStyleFrom(estiloSoma);
            cellSoma.setCellStyle(somaStyle);
        } else {
            CellStyle defaultStyle = workbook.createCellStyle();
            defaultStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
            cellSoma.setCellStyle(defaultStyle);
        }
    }
}
