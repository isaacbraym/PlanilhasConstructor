package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.*;
import java.util.*;

public class ManipuladorPlanilha {
    private Sheet sheet;
    private final LogsDeModificadores logs;
    private Map<Integer, CellData> colunaTemporaria = new HashMap<>();
    private int columnOffset;

    // Construtor que determina o columnOffset automaticamente
    public ManipuladorPlanilha(Sheet sheet) {
        this(sheet, determinarColunaInicial(sheet));
    }

    // Construtor que permite configurar o columnOffset manualmente
    public ManipuladorPlanilha(Sheet sheet, int columnOffset) {
        this.sheet = sheet;
        this.columnOffset = columnOffset;
        this.logs = new LogsDeModificadores();
    }

    // Setter para ajustar o offset se necessário
    public void setColumnOffset(int columnOffset) {
        this.columnOffset = columnOffset;
    }

    // Método para determinar dinamicamente o offset da coluna
    private static int determinarColunaInicial(Sheet sheet) {
        Row primeiraLinha = sheet.getRow(0);
        if (primeiraLinha != null) {
            for (Cell cell : primeiraLinha) {
                if (cell != null) {
                    String cellValue = getCellValueAsStringStatic(cell);
                    if (cellValue != null && !cellValue.trim().isEmpty()) {
                        return cell.getColumnIndex();
                    }
                }
            }
        }
        return 0; // Padrão para a primeira coluna se nenhum cabeçalho for encontrado
    }

    // Método para mover uma coluna
    public ManipuladorPlanilha moverColuna(String moverAColuna, String paraAPosicao) {
        int colunaOrigem = PosicaoConverter.converterColuna(moverAColuna) - columnOffset;
        int colunaDestino = PosicaoConverter.converterColuna(paraAPosicao) - columnOffset;

        if (colunaOrigem == colunaDestino) {
            return this;
        }

        Map<Integer, String> headerMap = getHeaderMap();
        String headerOrigem = headerMap.get(colunaOrigem);
        LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerOrigem, 
                PosicaoConverter.converterIndice(colunaOrigem + columnOffset), 
                PosicaoConverter.converterIndice(colunaDestino + columnOffset));
        LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Deslocamento de colunas", mainMovement);

        copiarColuna(colunaOrigem);

        if (colunaOrigem < colunaDestino) {
            shiftColumnsLeft(colunaOrigem + 1, colunaDestino);
            registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
        } else {
            shiftColumnsRight(colunaDestino, colunaOrigem - 1);
            registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
        }

        colarColunaTemporaria(colunaDestino);
        colunaTemporaria.clear();
        logs.adicionarLog(actionLog);

        return this;
    }

    // Método para remover uma coluna
    public ManipuladorPlanilha removerColuna(String coluna) {
        int colIndex = PosicaoConverter.converterColuna(coluna) - columnOffset;
        int lastColumn = getLastColumnNum();

        Map<Integer, String> headerMap = getHeaderMap();
        String headerName = headerMap.get(colIndex);
        LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerName, 
                PosicaoConverter.converterIndice(colIndex + columnOffset), null);
        LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Remoção de coluna", mainMovement);

        removerCelulasDaColuna(colIndex);
        if (colIndex < lastColumn) {
            shiftColumnsLeft(colIndex + 1, lastColumn);
            registrarColunasDeslocadasRemocao(colIndex, lastColumn, headerMap, actionLog);
        }

        logs.adicionarLog(actionLog);
        return this;
    }

    // Método para inserir uma coluna vazia entre duas colunas especificadas
    public ManipuladorPlanilha inserirColunaVaziaEntre(String colunaEsquerda, String colunaDireita) {
        int colEsquerdaIndex = PosicaoConverter.converterColuna(colunaEsquerda) - columnOffset;
        int colDireitaIndex = PosicaoConverter.converterColuna(colunaDireita) - columnOffset;

        validarAdjacencia(colEsquerdaIndex, colDireitaIndex, colunaEsquerda, colunaDireita);

        Map<Integer, String> headerMap = getHeaderMap();
        int posicaoInsercao = colDireitaIndex;
        int lastColumn = getLastColumnNum();

        LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Inserção de coluna vazia",
                new LogsDeModificadores.ColumnMovement(null, colunaEsquerda, colunaDireita));

        if (posicaoInsercao <= lastColumn) {
            shiftColumnsRight(posicaoInsercao, lastColumn);
            registrarColunasDeslocadasInsercao(posicaoInsercao, lastColumn, headerMap, actionLog);
            logs.adicionarLog(actionLog);
        }

        definirLarguraDaNovaColuna(posicaoInsercao);
        return this;
    }

    // Método para exibir o log de rastreio
    public void logAlteracoes() {
        logs.exibirLogs();
    }

    // Métodos auxiliares privados

    private void validarAdjacencia(int colEsquerdaIndex, int colDireitaIndex, String colunaEsquerda, String colunaDireita) {
        if (colDireitaIndex - colEsquerdaIndex != 1) {
            throw new IllegalArgumentException("As colunas especificadas não são adjacentes. Certifique-se de que "
                    + colunaDireita + " está imediatamente à direita de " + colunaEsquerda + ".");
        }
    }

    private void definirLarguraDaNovaColuna(int posicaoInsercao) {
        sheet.setColumnWidth(posicaoInsercao + columnOffset, sheet.getDefaultColumnWidth() * 256);
    }

    private void registrarColunasDeslocadas(int colunaOrigem, int colunaDestino, Map<Integer, String> headerMap, 
                                           LogsDeModificadores.ActionLog actionLog) {
        if (colunaOrigem < colunaDestino) {
            for (int col = colunaOrigem + 1; col <= colunaDestino; col++) {
                adicionarColunaDeslocada(col, col - 1, headerMap, actionLog);
            }
        } else {
            for (int col = colunaOrigem - 1; col >= colunaDestino; col--) {
                adicionarColunaDeslocada(col, col + 1, headerMap, actionLog);
            }
        }
    }

    private void registrarColunasDeslocadasRemocao(int colIndex, int lastColumn, Map<Integer, String> headerMap, 
                                                   LogsDeModificadores.ActionLog actionLog) {
        for (int col = colIndex + 1; col <= lastColumn; col++) {
            adicionarColunaDeslocada(col, col - 1, headerMap, actionLog);
        }
    }

    private void registrarColunasDeslocadasInsercao(int posicaoInsercao, int lastColumn, Map<Integer, String> headerMap, 
                                                    LogsDeModificadores.ActionLog actionLog) {
        for (int col = lastColumn; col >= posicaoInsercao; col--) {
            adicionarColunaDeslocada(col, col + 1, headerMap, actionLog);
        }
    }

    private void adicionarColunaDeslocada(int col, int targetCol, Map<Integer, String> headerMap, 
                                          LogsDeModificadores.ActionLog actionLog) {
        String headerName = headerMap.get(col);
        if (headerName != null && !headerName.trim().isEmpty()) {
            String previousIndex = PosicaoConverter.converterIndice(col + columnOffset);
            String newIndex = PosicaoConverter.converterIndice(targetCol + columnOffset);
            actionLog.getShiftedColumns().add(new LogsDeModificadores.ColumnMovement(headerName, previousIndex, newIndex));
        }
    }

    private void removerCelulasDaColuna(int colIndex) {
        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                Cell cellToRemove = row.getCell(colIndex + columnOffset);
                if (cellToRemove != null) {
                    row.removeCell(cellToRemove);
                }
            }
        }
    }

    // Método auxiliar para obter um mapa dos índices de colunas para nomes de cabeçalhos
    private Map<Integer, String> getHeaderMap() {
        Map<Integer, String> headerMap = new HashMap<>();
        int lastRowNum = sheet.getLastRowNum();
        int lastColNum = getLastColumnNum();

        for (int colIndex = 0; colIndex <= lastColNum; colIndex++) {
            String headerName = encontrarHeader(colIndex, lastRowNum);
            if (headerName != null) {
                headerMap.put(colIndex, headerName);
            }
        }
        return headerMap;
    }

    private String encontrarHeader(int colIndex, int lastRowNum) {
        for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                Cell cell = row.getCell(colIndex + columnOffset);
                if (cell != null) {
                    String cellValue = getCellValueAsStringStatic(cell);
                    if (cellValue != null && !cellValue.trim().isEmpty()) {
                        return cellValue;
                    }
                }
            }
        }
        return null;
    }

    // Método auxiliar estático para obter o valor da célula como String
    private static String getCellValueAsStringStatic(Cell cell) {
        if (cell == null)
            return null;

        CellType cellType = cell.getCellTypeEnum();
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return Double.toString(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case ERROR:
                return Byte.toString(cell.getErrorCellValue());
            default:
                return null;
        }
    }

    // Método auxiliar para obter o valor da célula como String
    private String getCellValueAsString(Cell cell) {
        return getCellValueAsStringStatic(cell);
    }

    // Método auxiliar para obter o índice da última coluna na planilha
    private int getLastColumnNum() {
        int lastCol = 0;
        for (Row row : sheet) {
            if (row.getLastCellNum() > lastCol) {
                lastCol = row.getLastCellNum();
            }
        }
        return lastCol - columnOffset - 1; // Ajuste pelo deslocamento e 1-based
    }

    // Método para copiar uma coluna para a coluna temporária
    private void copiarColuna(int colunaOrigem) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(colunaOrigem + columnOffset);
                if (cell != null) {
                    CellData cellData = new CellData();
                    copiarValorParaCellData(cell, cellData);
                    copiarEstiloParaCellData(cell, cellData);
                    colunaTemporaria.put(i, cellData);
                    row.removeCell(cell);
                }
            }
        }
    }

    private void copiarValorParaCellData(Cell cell, CellData cellData) {
        CellType cellType = cell.getCellTypeEnum();
        cellData.setCellType(cellType);

        switch (cellType) {
            case STRING:
                cellData.setStringValue(cell.getStringCellValue());
                break;
            case NUMERIC:
                cellData.setNumericValue(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                cellData.setBooleanValue(cell.getBooleanCellValue());
                break;
            case FORMULA:
                cellData.setFormulaValue(cell.getCellFormula());
                break;
            case ERROR:
                cellData.setErrorValue(cell.getErrorCellValue());
                break;
            case BLANK:
                // Nada a fazer para células em branco
                break;
            default:
                // Outros tipos de célula, se necessário
                break;
        }
    }

    private void copiarEstiloParaCellData(Cell cell, CellData cellData) {
        cellData.setCellStyle(cell.getCellStyle());
    }

    // Método para colar a coluna temporária na nova posição
    private void colarColunaTemporaria(int colunaDestino) {
        for (Map.Entry<Integer, CellData> entry : colunaTemporaria.entrySet()) {
            int rowNum = entry.getKey();
            CellData cellData = entry.getValue();

            Row row = sheet.getRow(rowNum);
            if (row == null) {
                row = sheet.createRow(rowNum);
            }
            Cell cell = row.createCell(colunaDestino + columnOffset);
            colarValorDeCellData(cell, cellData);
            cell.setCellStyle(cellData.getCellStyle());
        }
    }

    private void colarValorDeCellData(Cell cell, CellData cellData) {
        cell.setCellType(cellData.getCellType());

        switch (cellData.getCellType()) {
            case STRING:
                cell.setCellValue(cellData.getStringValue());
                break;
            case NUMERIC:
                cell.setCellValue(cellData.getNumericValue());
                break;
            case BOOLEAN:
                cell.setCellValue(cellData.isBooleanValue());
                break;
            case FORMULA:
                cell.setCellFormula(cellData.getFormulaValue());
                break;
            case ERROR:
                cell.setCellErrorValue(cellData.getErrorValue());
                break;
            case BLANK:
                // Deixar a célula em branco
                break;
            default:
                // Outros tipos de célula, se necessário
                break;
        }
    }

    // Métodos para deslocar colunas manualmente

    private void shiftColumnsLeft(int startColumn, int endColumn) {
        int lastRowNum = sheet.getLastRowNum();
        for (int col = startColumn; col <= endColumn; col++) {
            int targetCol = col - 1;
            for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    Cell sourceCell = row.getCell(col + columnOffset);
                    if (sourceCell != null) {
                        Cell targetCell = row.createCell(targetCol + columnOffset);
                        copiarValorEntreCelulas(sourceCell, targetCell);
                        copiarEstiloEntreCelulas(sourceCell, targetCell);
                        row.removeCell(sourceCell);
                    } else {
                        removerCelula(row, targetCol + columnOffset);
                    }
                }
            }
        }
    }

    private void shiftColumnsRight(int startColumn, int endColumn) {
        int lastRowNum = sheet.getLastRowNum();
        for (int col = endColumn; col >= startColumn; col--) {
            int targetCol = col + 1;
            for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    Cell sourceCell = row.getCell(col + columnOffset);
                    if (sourceCell != null) {
                        Cell targetCell = row.createCell(targetCol + columnOffset);
                        copiarValorEntreCelulas(sourceCell, targetCell);
                        copiarEstiloEntreCelulas(sourceCell, targetCell);
                        row.removeCell(sourceCell);
                    } else {
                        removerCelula(row, targetCol + columnOffset);
                    }
                }
            }
        }
    }

    private void removerCelula(Row row, int cellIndex) {
        Cell targetCell = row.getCell(cellIndex);
        if (targetCell != null) {
            row.removeCell(targetCell);
        }
    }

    // Métodos para copiar valores e estilos entre células

    private void copiarValorEntreCelulas(Cell sourceCell, Cell targetCell) {
        CellType cellType = sourceCell.getCellTypeEnum();
        targetCell.setCellType(cellType);

        switch (cellType) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case ERROR:
                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
            case BLANK:
                // Deixar a célula em branco
                break;
            default:
                // Outros tipos de célula, se necessário
                break;
        }
    }

    private void copiarEstiloEntreCelulas(Cell sourceCell, Cell targetCell) {
        targetCell.setCellStyle(sourceCell.getCellStyle());
    }

    // Classe auxiliar para armazenar os dados da célula
    private static class CellData {
        private String stringValue;
        private String formulaValue;
        private double numericValue;
        private boolean booleanValue;
        private byte errorValue;
        private CellType cellType;
        private CellStyle cellStyle;

        // Getters e Setters
        public String getStringValue() {
            return stringValue;
        }

        public void setStringValue(String stringValue) {
            this.stringValue = stringValue;
        }

        public String getFormulaValue() {
            return formulaValue;
        }

        public void setFormulaValue(String formulaValue) {
            this.formulaValue = formulaValue;
        }

        public double getNumericValue() {
            return numericValue;
        }

        public void setNumericValue(double numericValue) {
            this.numericValue = numericValue;
        }

        public boolean isBooleanValue() {
            return booleanValue;
        }

        public void setBooleanValue(boolean booleanValue) {
            this.booleanValue = booleanValue;
        }

        public byte getErrorValue() {
            return errorValue;
        }

        public void setErrorValue(byte errorValue) {
            this.errorValue = errorValue;
        }

        public CellType getCellType() {
            return cellType;
        }

        public void setCellType(CellType cellType) {
            this.cellType = cellType;
        }

        public CellStyle getCellStyle() {
            return cellStyle;
        }

        public void setCellStyle(CellStyle cellStyle) {
            this.cellStyle = cellStyle;
        }
    }
}
