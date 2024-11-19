package com.abnote.planilhas.estilos;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.estilos.BackGroundColor;
import com.abnote.planilhas.estilos.estilos.BoldStyle;
import com.abnote.planilhas.estilos.estilos.BorderStyleHelper;
import com.abnote.planilhas.estilos.estilos.CenterStyle;
import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.estilos.estilos.FonteEnum;
import com.abnote.planilhas.estilos.estilos.Fontes;

public class EstiloCelula {
	private final Workbook workbook;
	private final Sheet sheet;

	private final int rowIndex;
	private final int columnIndex;

	private final int startRowIndex;
	private final int startColumnIndex;
	private final int endRowIndex;
	private final int endColumnIndex;

	private final boolean isRange;

	private final Map<String, org.apache.poi.ss.usermodel.CellStyle> styleCache = new HashMap<>();

	// Construtor para aplicar estilos na planilha inteira
	public EstiloCelula(Workbook workbook, Sheet sheet) {
		this(workbook, sheet, -1, -1, 0, 0, sheet.getLastRowNum(), getMaxColumnIndex(sheet));
	}

	// Construtor para aplicar estilos em uma célula específica
	public EstiloCelula(Workbook workbook, Sheet sheet, int rowIndex, int columnIndex) {
		this(workbook, sheet, rowIndex, columnIndex, -1, -1, -1, -1);
	}

	// Construtor para aplicar estilos em um intervalo
	public EstiloCelula(Workbook workbook, Sheet sheet, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex) {
		this(workbook, sheet, startRowIndex, startColumnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex);
	}

	// Construtor interno que inicializa todos os campos
	private EstiloCelula(Workbook workbook, Sheet sheet, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex) {
		this.workbook = workbook;
		this.sheet = sheet;

		this.rowIndex = rowIndex;
		this.columnIndex = columnIndex;

		this.startRowIndex = startRowIndex;
		this.startColumnIndex = startColumnIndex;

		this.endRowIndex = endRowIndex;
		this.endColumnIndex = endColumnIndex;

		this.isRange = (endRowIndex != -1 && endColumnIndex != -1);
	}

	// Método auxiliar para obter o maior índice de coluna na planilha
	private static int getMaxColumnIndex(Sheet sheet) {
		int maxColIndex = -1;
		for (int rowIdx = 0; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			if (row != null && row.getLastCellNum() > maxColIndex) {
				maxColIndex = row.getLastCellNum();
			}
		}
		return maxColIndex - 1; // Ajuste para índice baseado em zero
	}

	// Métodos para aplicar estilos

	public EstiloCelula aplicarBold() {
		BoldStyle boldStyle = new BoldStyle(workbook, sheet, styleCache);
		boldStyle.aplicarBold(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}

	public EstiloCelula todasAsBordasEmTudo() {
		BorderStyleHelper borderStyleHelper = new BorderStyleHelper(workbook, sheet, styleCache);
		borderStyleHelper.todasAsBordasEmTudo(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula aplicarBordasNaCelula(String posicao) {
		BorderStyleHelper borderStyleHelper = new BorderStyleHelper(workbook, sheet, styleCache);
		borderStyleHelper.aplicarBordasNaCelula(posicao);
		return this;
	}

	public EstiloCelula aplicarTodasAsBordasDeAte(String posicaoInicial, String posicaoFinal) {
		BorderStyleHelper borderStyleHelper = new BorderStyleHelper(workbook, sheet, styleCache);
		borderStyleHelper.aplicarTodasAsBordasDeAte(posicaoInicial, posicaoFinal);
		return this;
	}

	public EstiloCelula bordasEspessas(String posicaoInicial, String posicaoFinal) {
		BorderStyleHelper borderStyleHelper = new BorderStyleHelper(workbook, sheet, styleCache);
		borderStyleHelper.bordasEspessas(posicaoInicial, posicaoFinal);
		return this;
	}

	public EstiloCelula bordasEspessasComBordasInternas(String posicaoInicial, String posicaoFinal) {
		BorderStyleHelper borderStyleHelper = new BorderStyleHelper(workbook, sheet, styleCache);
		borderStyleHelper.bordasEspessasComBordasInternas(posicaoInicial, posicaoFinal);
		return this;
	}

	public EstiloCelula centralizarTudo() {
		CenterStyle centerStyle = new CenterStyle(workbook, sheet, styleCache);
		centerStyle.centralizarTudo(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula centralizarERedimensionarTudo() {
		CenterStyle centerStyle = new CenterStyle(workbook, sheet, styleCache);
		centerStyle.centralizarERedimensionarTudo();
		return this;
	}

	public EstiloCelula redimensionarColunas() {
		CenterStyle centerStyle = new CenterStyle(workbook, sheet, styleCache);
		centerStyle.redimensionarColunas();
		return this;
	}

	public EstiloCelula removerLinhasDeGrade() {
		sheet.setDisplayGridlines(false);
		return this;
	}

	// Métodos para aplicar estilos de fonte

	public EstiloCelula fonte(String fontName) {
		Fontes fontes = new Fontes(workbook, sheet, styleCache);
		fontes.aplicarFonte(fontName, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula fonte(FonteEnum fonteEnum) {
		Fontes fontes = new Fontes(workbook, sheet, styleCache);
		fontes.aplicarFonte(fonteEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula fonteTamanho(int fontSize) {
		Fontes fontes = new Fontes(workbook, sheet, styleCache);
		fontes.aplicarTamanhoFonte(fontSize, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corFonte(CorEnum corEnum) {
		Fontes fontes = new Fontes(workbook, sheet, styleCache);
		fontes.aplicarCorFonte(corEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corFonte(int red, int green, int blue) {
		Fontes fontes = new Fontes(workbook, sheet, styleCache);
		fontes.aplicarCorFonte(red, green, blue, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corFonte(String hexColor) {
		Fontes fontes = new Fontes(workbook, sheet, styleCache);
		fontes.aplicarCorFonte(hexColor, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	// Métodos para aplicar cor de fundo

	public EstiloCelula corDeFundo(CorEnum corEnum) {
		BackGroundColor bgColor = new BackGroundColor(workbook, sheet, styleCache);
		bgColor.aplicarCorDeFundo(corEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corDeFundo(int red, int green, int blue) {
		BackGroundColor bgColor = new BackGroundColor(workbook, sheet, styleCache);
		bgColor.aplicarCorDeFundo(red, green, blue, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	public EstiloCelula corDeFundo(String hexColor) {
		BackGroundColor bgColor = new BackGroundColor(workbook, sheet, styleCache);
		bgColor.aplicarCorDeFundo(hexColor, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	// Getters para uso nas classes auxiliares

	public Workbook getWorkbook() {
		return workbook;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	public int getStartRowIndex() {
		return startRowIndex;
	}

	public int getStartColumnIndex() {
		return startColumnIndex;
	}

	public int getEndRowIndex() {
		return endRowIndex;
	}

	public int getEndColumnIndex() {
		return endColumnIndex;
	}

	public boolean isRange() {
		return isRange;
	}

	public Map<String, org.apache.poi.ss.usermodel.CellStyle> getStyleCache() {
		return styleCache;
	}
}
