package com.abnote.planilhas.estilos.estilos;

import java.util.Map;
import java.util.Objects;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * Classe responsável por aplicar estilos de negrito em células, linhas ou
 * intervalos de uma planilha.
 */
public class BoldStyle {

	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	/**
	 * Construtor para inicializar o BoldStyle com um Workbook, Sheet e cache de
	 * estilos.
	 *
	 * @param workbook   O Workbook que contém a planilha.
	 * @param sheet      A Sheet onde os estilos serão aplicados.
	 * @param styleCache Cache para armazenar estilos já criados.
	 */
	public BoldStyle(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	/**
	 * Aplica negrito em uma célula específica, em uma linha inteira ou em um
	 * intervalo de células.
	 *
	 * @param rowIndex         Índice da linha (-1 para não especificar).
	 * @param columnIndex      Índice da coluna (-1 para não especificar).
	 * @param startRowIndex    Índice da primeira linha do intervalo.
	 * @param startColumnIndex Índice da primeira coluna do intervalo.
	 * @param endRowIndex      Índice da última linha do intervalo.
	 * @param endColumnIndex   Índice da última coluna do intervalo.
	 * @param isRange          Se verdadeiro, aplica em um intervalo; caso
	 *                         contrário, aplica em linha ou célula específica.
	 */
	public void aplicarNegrito(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		if (isRange) {
			aplicarNegritoEmIntervalo(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
		} else if (rowIndex != -1) {
			if (columnIndex == -1) {
				aplicarNegritoEmLinha(rowIndex);
			} else {
				aplicarNegritoEmCelulaEspecifica(rowIndex, columnIndex);
			}
		}
	}

	// Método privado para aplicar negrito em um intervalo específico de células
	private void aplicarNegritoEmIntervalo(int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex) {
		for (int rowIdx = startRowIndex; rowIdx <= endRowIndex; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row == null)
				continue;

			for (int colIdx = startColumnIndex; colIdx <= endColumnIndex; colIdx++) {
				Cell cell = row.getCell(colIdx);
				if (cell == null)
					continue;
				aplicarNegritoNaCelula(cell);
			}
		}
	}

	// Método privado para aplicar negrito em toda uma linha
	private void aplicarNegritoEmLinha(int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row != null) {
			for (Cell cell : row) {
				if (cell != null) {
					aplicarNegritoNaCelula(cell);
				}
			}
		}
	}

	// Método privado para aplicar negrito em uma célula específica
	private void aplicarNegritoEmCelulaEspecifica(int rowIndex, int columnIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row != null) {
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				aplicarNegritoNaCelula(cell);
			}
		}
	}

	// Método privado para aplicar negrito em uma célula específica
	private void aplicarNegritoNaCelula(Cell cell) {
		CellStyle estiloAtual = cell.getCellStyle();
		Font fonteAtual = workbook.getFontAt(estiloAtual.getFontIndex());

		Font fonteNegrito = obterOuCriarFonteNegrito(fonteAtual);

		String chaveEstilo = "bold_" + estiloAtual.hashCode();
		CellStyle novoEstilo = styleCache.get(chaveEstilo);
		if (novoEstilo == null) {
			novoEstilo = workbook.createCellStyle();
			novoEstilo.cloneStyleFrom(estiloAtual);
			novoEstilo.setFont(fonteNegrito);
			styleCache.put(chaveEstilo, novoEstilo);
		}

		cell.setCellStyle(novoEstilo);
	}

	// Método privado para obter uma fonte negrito existente ou criar uma nova
	private Font obterOuCriarFonteNegrito(Font fonteAtual) {
		Font fonteNegritoExistente = buscarFonteNegritoExistente(fonteAtual);
		if (fonteNegritoExistente != null) {
			return fonteNegritoExistente;
		} else {
			return criarFonteNegrito(fonteAtual);
		}
	}

	// Método privado para buscar uma fonte negrito existente que corresponda à
	// fonte atual (exceto negrito)
	private Font buscarFonteNegritoExistente(Font fonteAtual) {
		for (short i = 0; i < workbook.getNumberOfFonts(); i++) {
			Font fonte = workbook.getFontAt(i);
			if (fonte.getBold() && fontesSaoIguaisExcetoNegrito(fonte, fonteAtual)) {
				return fonte;
			}
		}
		return null;
	}

	// Método privado para criar uma nova fonte negrito baseada na fonte atual
	private Font criarFonteNegrito(Font fonteAtual) {
		Font novaFonte = workbook.createFont();
		copiarAtributosFonte(fonteAtual, novaFonte);
		novaFonte.setBold(true);
		return novaFonte;
	}

	// Método privado para verificar se duas fontes são iguais exceto pelo atributo
	// negrito
	private boolean fontesSaoIguaisExcetoNegrito(Font fonte1, Font fonte2) {
		return fonte1.getFontName().equals(fonte2.getFontName())
				&& fonte1.getFontHeightInPoints() == fonte2.getFontHeightInPoints()
				&& fonte1.getItalic() == fonte2.getItalic() && fonte1.getStrikeout() == fonte2.getStrikeout()
				&& fonte1.getTypeOffset() == fonte2.getTypeOffset() && fonte1.getUnderline() == fonte2.getUnderline()
				&& fonte1.getCharSet() == fonte2.getCharSet() && fontesSaoComAMesmaCor(fonte1, fonte2);
	}

	// Método privado para copiar atributos de uma fonte para outra
	private void copiarAtributosFonte(Font fonteOrigem, Font fonteDestino) {
		fonteDestino.setFontName(fonteOrigem.getFontName());
		fonteDestino.setFontHeightInPoints(fonteOrigem.getFontHeightInPoints());
		fonteDestino.setItalic(fonteOrigem.getItalic());
		fonteDestino.setStrikeout(fonteOrigem.getStrikeout());
		fonteDestino.setTypeOffset(fonteOrigem.getTypeOffset());
		fonteDestino.setUnderline(fonteOrigem.getUnderline());
		fonteDestino.setCharSet(fonteOrigem.getCharSet());
		fonteDestino.setColor(fonteOrigem.getColor());

		if (fonteOrigem instanceof XSSFFont && fonteDestino instanceof XSSFFont) {
			XSSFColor cor = ((XSSFFont) fonteOrigem).getXSSFColor();
			((XSSFFont) fonteDestino).setColor(cor);
		}
	}

	// Método privado para verificar se duas fontes têm a mesma cor
	private boolean fontesSaoComAMesmaCor(Font fonte1, Font fonte2) {
		if (fonte1 instanceof XSSFFont && fonte2 instanceof XSSFFont) {
			XSSFColor cor1 = ((XSSFFont) fonte1).getXSSFColor();
			XSSFColor cor2 = ((XSSFFont) fonte2).getXSSFColor();
			return Objects.equals(cor1, cor2);
		} else {
			return fonte1.getColor() == fonte2.getColor();
		}
	}
}
