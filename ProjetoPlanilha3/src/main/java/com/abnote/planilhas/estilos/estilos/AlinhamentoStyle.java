package com.abnote.planilhas.estilos.estilos;

import java.util.Map;

import javax.swing.GroupLayout.Alignment;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class AlinhamentoStyle {
	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	public AlinhamentoStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	// Método genérico para aplicar alinhamento e quebra de texto
	private void aplicarAlinhamento(HorizontalAlignment alignment, boolean quebraTexto, int rowIndex, int columnIndex,
			int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		iterateCells(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange,
				(Cell cell) -> applyAlignmentToCell(cell, alignment, quebraTexto));
	}

	public void alinharAEsquerda(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		aplicarAlinhamento(HorizontalAlignment.LEFT, false, rowIndex, columnIndex, startRowIndex, startColumnIndex,
				endRowIndex, endColumnIndex, isRange);
	}

	public void alinharADireita(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		aplicarAlinhamento(HorizontalAlignment.RIGHT, false, rowIndex, columnIndex, startRowIndex, startColumnIndex,
				endRowIndex, endColumnIndex, isRange);
	}

	public void quebrarTexto(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		aplicarAlinhamento(null, true, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	// Método para iterar sobre as células e aplicar uma ação
	private void iterateCells(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange, CellAction action) {

		if (isRange) {
			for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
				Row row = sheet.getRow(rowIdx);
				if (row == null)
					continue;
				for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
					Cell cell = row.getCell(colIdx);
					if (cell != null) {
						action.apply(cell);
					}
				}
			}
		} else if (rowIndex != -1) {
			if (columnIndex == -1) {
				// Aplicar à linha inteira
				Row row = sheet.getRow(rowIndex);
				if (row != null) {
					for (Cell cell : row) {
						if (cell != null) {
							action.apply(cell);
						}
					}
				}
			} else {
				// Aplicar à célula específica
				Row row = sheet.getRow(rowIndex);
				if (row != null) {
					Cell cell = row.getCell(columnIndex);
					if (cell != null) {
						action.apply(cell);
					}
				}
			}
		}
	}

	// Interface funcional para aplicar ações nas células
	@FunctionalInterface
	private interface CellAction {
		void apply(Cell cell);
	}

	// Método para aplicar alinhamento e quebra de texto a uma célula
	private void applyAlignmentToCell(Cell cell, HorizontalAlignment alignment, boolean quebraTexto) {
		CellStyle currentStyle = cell.getCellStyle();

		// Geração de uma chave única para o cache com base no estilo atual, alinhamento
		// e quebra de texto
		String key = "alignment_" + currentStyle.hashCode() + "_" + (alignment != null ? alignment.name() : "null")
				+ "_" + quebraTexto;

		CellStyle newStyle = styleCache.get(key);
		if (newStyle == null) {
			newStyle = workbook.createCellStyle();
			newStyle.cloneStyleFrom(currentStyle);

			if (alignment != null) {
				newStyle.setAlignment(alignment);
			}

			newStyle.setWrapText(quebraTexto);

			styleCache.put(key, newStyle);
		}

		cell.setCellStyle(newStyle);
	}
}
