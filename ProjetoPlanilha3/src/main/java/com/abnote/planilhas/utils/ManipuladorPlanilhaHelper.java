package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import java.util.*;

public class ManipuladorPlanilhaHelper {
	private Sheet sheet;
	private int columnOffset;

	// Construtor
	public ManipuladorPlanilhaHelper(Sheet sheet, int columnOffset) {
		this.sheet = sheet;
		this.columnOffset = columnOffset;
	}

	// Método para determinar dinamicamente o offset da coluna
	public static int determinarColunaInicial(Sheet sheet) {
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

	/**
	 * Limpa os dados de uma coluna específica sem remover ou deslocar a coluna.
	 *
	 * @param colIndex Índice da coluna a ser limpa
	 */
	public void limparColuna(int colIndex) {
		int lastRowNum = sheet.getLastRowNum();
		for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row != null) {
				Cell cell = row.getCell(colIndex + columnOffset);
				if (cell != null) {
					cell.setCellType(CellType.BLANK);
				}
			}
		}
	}

	// Método auxiliar estático para obter o valor da célula como String
	public static String getCellValueAsStringStatic(Cell cell) {
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

	// Método auxiliar para obter o índice da última coluna na planilha
	public int getLastColumnNum() {
		int lastCol = 0;
		for (Row row : sheet) {
			if (row.getLastCellNum() > lastCol) {
				lastCol = row.getLastCellNum();
			}
		}
		return lastCol - columnOffset - 1; // Ajuste pelo deslocamento e 1-based
	}

	// Método auxiliar para obter um mapa dos índices de colunas para nomes de
	// cabeçalhos
	public Map<Integer, String> getHeaderMap() {
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

	// Valida se as colunas são adjacentes
	public void validarAdjacencia(int colEsquerdaIndex, int colDireitaIndex, String colunaEsquerda,
			String colunaDireita) {
		if (colDireitaIndex - colEsquerdaIndex != 1) {
			throw new IllegalArgumentException("As colunas especificadas não são adjacentes. Certifique-se de que "
					+ colunaDireita + " está imediatamente à direita de " + colunaEsquerda + ".");
		}
	}

	// Define a largura da nova coluna
	public void definirLarguraDaNovaColuna(int posicaoInsercao) {
		sheet.setColumnWidth(posicaoInsercao + columnOffset, sheet.getDefaultColumnWidth() * 256);
	}

	// Registrar colunas deslocadas após movimentação
	public void registrarColunasDeslocadas(int colunaOrigem, int colunaDestino, Map<Integer, String> headerMap,
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

	// Registrar colunas deslocadas após remoção
	public void registrarColunasDeslocadasRemocao(int colIndex, int lastColumn, Map<Integer, String> headerMap,
			LogsDeModificadores.ActionLog actionLog) {
		for (int col = colIndex + 1; col <= lastColumn; col++) {
			adicionarColunaDeslocada(col, col - 1, headerMap, actionLog);
		}
	}

	// Registrar colunas deslocadas após inserção
	public void registrarColunasDeslocadasInsercao(int posicaoInsercao, int lastColumn, Map<Integer, String> headerMap,
			LogsDeModificadores.ActionLog actionLog) {
		for (int col = lastColumn; col >= posicaoInsercao; col--) {
			adicionarColunaDeslocada(col, col + 1, headerMap, actionLog);
		}
	}

	// Adiciona uma coluna deslocada ao log
	private void adicionarColunaDeslocada(int col, int targetCol, Map<Integer, String> headerMap,
			LogsDeModificadores.ActionLog actionLog) {
		String headerName = headerMap.get(col);
		if (headerName != null && !headerName.trim().isEmpty()) {
			String previousIndex = PosicaoConverter.converterIndice(col + columnOffset);
			String newIndex = PosicaoConverter.converterIndice(targetCol + columnOffset);
			actionLog.getShiftedColumns()
					.add(new LogsDeModificadores.ColumnMovement(headerName, previousIndex, newIndex));
		}
	}

	// Remove as células de uma coluna
	public void removerCelulasDaColuna(int colIndex) {
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

	// Copiar uma coluna para a coluna temporária
	public Map<Integer, CellData> copiarColuna(int colunaOrigem) {
		Map<Integer, CellData> colunaTemporaria = new HashMap<>();
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
		return colunaTemporaria;
	}

	// Colar a coluna temporária na nova posição
	public void colarColunaTemporaria(int colunaDestino, Map<Integer, CellData> colunaTemporaria) {
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
	public void shiftColumnsLeft(int startColumn, int endColumn) {
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

	public void shiftColumnsRight(int startColumn, int endColumn) {
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
	public static class CellData {
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
