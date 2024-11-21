package com.abnote.planilhas.estilos.estilos;

import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class CenterStyle {
	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	public CenterStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	/**
	 * Centraliza as células em um intervalo específico ou em toda a planilha.
	 *
	 * @param startRow    Índice da primeira linha
	 * @param startColumn Índice da primeira coluna
	 * @param endRow      Índice da última linha
	 * @param endColumn   Índice da última coluna
	 * @param isRange     Se verdadeiro, aplica ao intervalo; caso contrário, a toda
	 *                    a planilha
	 */
	public void centralizarTudo(int startRow, int startColumn, int endRow, int endColumn, boolean isRange) {
		if (isRange) {
			centralizarIntervalo(startRow, startColumn, endRow, endColumn);
		} else {
			centralizarPlanilha();
		}
	}

	public void centralizarERedimensionarTudo() {
		centralizarPlanilha();
		redimensionarColunas();
	}

	private void centralizarIntervalo(int startRow, int startColumn, int endRow, int endColumn) {
		for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row == null)
				continue;

			for (int colIdx = startColumn; colIdx <= endColumn; colIdx++) {
				Cell cell = row.getCell(colIdx);
				if (cell == null)
					continue;

				aplicarCentralizacao(cell);
			}
		}
	}

	// Método privado para centralizar todas as células da planilha
	private void centralizarPlanilha() {
		for (Row row : sheet) {
			if (row == null)
				continue;

			for (Cell cell : row) {
				if (cell == null)
					continue;

				aplicarCentralizacao(cell);
			}
		}
	}

	// Método privado para aplicar a centralização a uma célula específica
	private void aplicarCentralizacao(Cell cell) {
		CellStyle originalStyle = cell.getCellStyle();
		CellStyle novoEstilo = criarEstiloCentralizado(originalStyle);
		cell.setCellStyle(novoEstilo);
	}

	// Método privado para criar um novo estilo com centralização
	private CellStyle criarEstiloCentralizado(CellStyle originalStyle) {
		CellStyle novoEstilo = workbook.createCellStyle();
		novoEstilo.cloneStyleFrom(originalStyle);
		novoEstilo.setAlignment(HorizontalAlignment.CENTER);
		novoEstilo.setVerticalAlignment(VerticalAlignment.CENTER);
		return novoEstilo;
	}

	// Método para redimensionar todas as colunas com base no conteúdo
	public void redimensionarColunas() {
		int maxColumns = obterMaximoNumeroDeColunas();
		for (int i = 0; i < maxColumns; i++) {
			sheet.autoSizeColumn(i);
		}
	}

	// Método privado para determinar o número máximo de colunas na planilha
	private int obterMaximoNumeroDeColunas() {
		int maxColumns = 0;
		for (Row row : sheet) {
			if (row == null)
				continue;

			int lastCellNum = row.getLastCellNum();
			if (lastCellNum > maxColumns) {
				maxColumns = lastCellNum;
			}
		}
		return maxColumns;
	}
}
