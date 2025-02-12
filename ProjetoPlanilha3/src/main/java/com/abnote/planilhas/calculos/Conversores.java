package com.abnote.planilhas.calculos;

import org.apache.poi.ss.usermodel.*;
import com.abnote.planilhas.utils.PosicaoConverter;

public class Conversores {

	/**
	 * Converte os valores de uma coluna para números, se possível.
	 *
	 * @param sheet          A folha da planilha a ser processada.
	 * @param posicaoInicial A posição inicial da coluna (ex: "J3").
	 */
	public static void converterEmNumero(Sheet sheet, String posicaoInicial) {
		int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
		int coluna = posicao[0];
		int linhaInicial = posicao[1];

		for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(coluna);
			if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
				try {
					double valorNumerico = Double.parseDouble(cell.getStringCellValue());
					cell.setCellType(CellType.NUMERIC);
					cell.setCellValue(valorNumerico);
				} catch (NumberFormatException e) {
					System.out.println("Célula em " + (i + 1) + " não é numérica e foi ignorada.");
				}
			}
		}
	}

	/**
	 * Converte os valores de uma coluna para o formato contábil.
	 *
	 * @param sheet          A folha da planilha a ser processada.
	 * @param posicaoInicial A posição inicial da coluna (ex: "J3").
	 * @param workbook       O workbook da planilha para criar estilos.
	 */
	public static void converterEmContabil(Sheet sheet, String posicaoInicial, Workbook workbook) {
		int[] posicao = PosicaoConverter.converterPosicao(posicaoInicial);
		int coluna = posicao[0];
		int linhaInicial = posicao[1];

		// Configuração do estilo contábil para Real (R$)
		CellStyle estiloContabil = workbook.createCellStyle();
		DataFormat formato = workbook.createDataFormat();
		estiloContabil.setDataFormat(formato.getFormat("#,##0.00"));

		for (int i = linhaInicial; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(coluna);
			if (cell != null) {
				if (cell.getCellTypeEnum() == CellType.STRING) {
					try {
						double valorNumerico = Double.parseDouble(cell.getStringCellValue());
						cell.setCellType(CellType.NUMERIC);
						cell.setCellValue(valorNumerico);
					} catch (NumberFormatException e) {
						System.out.println("Célula em " + (i + 1) + " não é numérica e foi ignorada.");
					}
				}
				if (cell.getCellTypeEnum() == CellType.NUMERIC) {
					cell.setCellStyle(estiloContabil);
				}
			}
		}
	}
}
