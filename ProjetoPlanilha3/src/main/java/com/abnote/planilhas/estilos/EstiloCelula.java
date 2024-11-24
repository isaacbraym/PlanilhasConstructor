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

/**
 * Classe responsável por aplicar diversos estilos em células, linhas ou
 * intervalos de uma planilha.
 */
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

	// Instâncias das classes auxiliares
	private final BoldStyle boldStyle;
	private final BorderStyleHelper borderStyleHelper;
	private final CenterStyle centerStyle;
	private final Fontes fontes;
	private final BackGroundColor backGroundColor;

	/**
	 * Construtor para aplicar estilos na planilha inteira.
	 *
	 * @param workbook O Workbook que contém a planilha.
	 * @param sheet    A Sheet onde os estilos serão aplicados.
	 */
	public EstiloCelula(Workbook workbook, Sheet sheet) {
		this(workbook, sheet, -1, -1, 0, 0, sheet.getLastRowNum(), getMaxColumnIndex(sheet));
	}

	/**
	 * Construtor para aplicar estilos em uma célula específica.
	 *
	 * @param workbook    O Workbook que contém a planilha.
	 * @param sheet       A Sheet onde os estilos serão aplicados.
	 * @param rowIndex    Índice da linha da célula (-1 para não especificar).
	 * @param columnIndex Índice da coluna da célula (-1 para não especificar).
	 */
	public EstiloCelula(Workbook workbook, Sheet sheet, int rowIndex, int columnIndex) {
		this(workbook, sheet, rowIndex, columnIndex, -1, -1, -1, -1);
	}

	/**
	 * Construtor para aplicar estilos em um intervalo.
	 *
	 * @param workbook         O Workbook que contém a planilha.
	 * @param sheet            A Sheet onde os estilos serão aplicados.
	 * @param startRowIndex    Índice da primeira linha do intervalo.
	 * @param startColumnIndex Índice da primeira coluna do intervalo.
	 * @param endRowIndex      Índice da última linha do intervalo.
	 * @param endColumnIndex   Índice da última coluna do intervalo.
	 */
	public EstiloCelula(Workbook workbook, Sheet sheet, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex) {
		this(workbook, sheet, startRowIndex, startColumnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex);
	}

	/**
	 * Construtor interno que inicializa todos os campos e as classes auxiliares.
	 *
	 * @param workbook         O Workbook que contém a planilha.
	 * @param sheet            A Sheet onde os estilos serão aplicados.
	 * @param rowIndex         Índice da linha da célula (-1 para não especificar).
	 * @param columnIndex      Índice da coluna da célula (-1 para não especificar).
	 * @param startRowIndex    Índice da primeira linha do intervalo.
	 * @param startColumnIndex Índice da primeira coluna do intervalo.
	 * @param endRowIndex      Índice da última linha do intervalo.
	 * @param endColumnIndex   Índice da última coluna do intervalo.
	 */
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

		// Inicialização das classes auxiliares com o cache de estilos
		this.boldStyle = new BoldStyle(workbook, sheet, styleCache);
		this.borderStyleHelper = new BorderStyleHelper(workbook, sheet, styleCache);
		this.centerStyle = new CenterStyle(workbook, sheet, styleCache);
		this.fontes = new Fontes(workbook, sheet, styleCache);
		this.backGroundColor = new BackGroundColor(workbook, sheet, styleCache);
	}

	/**
	 * Aplica o estilo itálico em uma célula específica, em uma linha inteira ou em
	 * um intervalo de células.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarItalico() {
		fontes.aplicarItalico(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}

	/**
	 * Aplica o estilo sublinhado em uma célula específica, em uma linha inteira ou
	 * em um intervalo de células.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarSublinhado() {
		fontes.aplicarSublinhado(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}
	
    public EstiloCelula aplicarTachado() {
        fontes.aplicarTachado(rowIndex, columnIndex, startRowIndex, startColumnIndex,
                             endRowIndex, endColumnIndex, isRange);
        return this;
	}

	/**
	 * Método auxiliar para obter o maior índice de coluna na planilha.
	 *
	 * @param sheet A Sheet para verificar.
	 * @return O maior índice de coluna encontrado.
	 */
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

	/**
	 * Aplica negrito em uma célula específica, em uma linha inteira ou em um
	 * intervalo de células.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarNegrito() {
		boldStyle.aplicarNegrito(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex,
				isRange);
		return this;
	}

	/**
	 * Aplica todas as bordas finas em um intervalo específico ou em toda a
	 * planilha.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarTodasAsBordas() {
		borderStyleHelper.aplicarTodasAsBordas(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica bordas finas em uma célula específica baseada na posição (e.g., "A1").
	 *
	 * @param posicao A posição da célula (ex: "A1").
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarBordasNaCelula(String posicao) {
		borderStyleHelper.aplicarBordasNaCelula(posicao);
		return this;
	}

	/**
	 * Aplica bordas finas entre duas posições específicas (e.g., "A1" até "C3").
	 *
	 * @param posicaoInicial A posição inicial (ex: "A1").
	 * @param posicaoFinal   A posição final (ex: "C3").
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarBordasEntre(String posicaoInicial, String posicaoFinal) {
		borderStyleHelper.aplicarBordasEntre(posicaoInicial, posicaoFinal);
		return this;
	}

	/**
	 * Aplica bordas finas nas bordas externas de um intervalo específico (e.g.,
	 * "A1" até "C3").
	 *
	 * @param posicaoInicial A posição inicial (ex: "A1").
	 * @param posicaoFinal   A posição final (ex: "C3").
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarBordasEspessas(String posicaoInicial, String posicaoFinal) {
		borderStyleHelper.aplicarBordasEspessas(posicaoInicial, posicaoFinal);
		return this;
	}

	/**
	 * Aplica bordas espessas nas bordas internas e externas de um intervalo
	 * específico (e.g., "A1" até "C3").
	 *
	 * @param posicaoInicial A posição inicial (ex: "A1").
	 * @param posicaoFinal   A posição final (ex: "C3").
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula aplicarBordasEspessasComInternas(String posicaoInicial, String posicaoFinal) {
		borderStyleHelper.aplicarBordasEspessasComInternas(posicaoInicial, posicaoFinal);
		return this;
	}

	/**
	 * Centraliza todas as células em um intervalo específico ou em toda a planilha.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula centralizarTudo() {
		centerStyle.centralizarTudo(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange);
		return this;
	}

	/**
	 * Centraliza todas as células e redimensiona as colunas com base no conteúdo.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula centralizarERedimensionarTudo() {
		centerStyle.centralizarERedimensionarTudo();
		return this;
	}

	/**
	 * Redimensiona todas as colunas com base no conteúdo.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula redimensionarColunas() {
		centerStyle.redimensionarColunas();
		return this;
	}

	/**
	 * Remove as linhas de grade da planilha.
	 *
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula removerLinhasDeGrade() {
		sheet.setDisplayGridlines(false);
		return this;
	}

	// Métodos para aplicar estilos de fonte

	/**
	 * Aplica uma fonte específica em uma célula, linha ou intervalo.
	 *
	 * @param fontName O nome da fonte a ser aplicada.
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula fonte(String fontName) {
		fontes.aplicarFonte(fontName, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica uma fonte específica em uma célula, linha ou intervalo.
	 *
	 * @param fonteEnum O enum da fonte a ser aplicada.
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula fonte(FonteEnum fonteEnum) {
		fontes.aplicarFonte(fonteEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica um tamanho de fonte específico em uma célula, linha ou intervalo.
	 *
	 * @param fontSize O tamanho da fonte a ser aplicada.
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula fonteTamanho(int fontSize) {
		fontes.aplicarTamanhoFonte(fontSize, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica uma cor específica na fonte de uma célula, linha ou intervalo.
	 *
	 * @param corEnum O enum da cor a ser aplicada.
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula corFonte(CorEnum corEnum) {
		fontes.aplicarCorFonte(corEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica uma cor específica na fonte de uma célula, linha ou intervalo usando
	 * valores RGB.
	 *
	 * @param red   Valor de vermelho (0-255).
	 * @param green Valor de verde (0-255).
	 * @param blue  Valor de azul (0-255).
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula corFonte(int red, int green, int blue) {
		fontes.aplicarCorFonte(red, green, blue, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica uma cor específica na fonte de uma célula, linha ou intervalo usando
	 * um código hexadecimal.
	 *
	 * @param hexColor O código hexadecimal da cor (ex: "#FFFFFF").
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula corFonte(String hexColor) {
		fontes.aplicarCorFonte(hexColor, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	// Métodos para aplicar cor de fundo

	/**
	 * Aplica uma cor de fundo específica em uma célula, linha ou intervalo.
	 *
	 * @param corEnum O enum da cor de fundo a ser aplicada.
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula corDeFundo(CorEnum corEnum) {
		backGroundColor.aplicarCorDeFundo(corEnum, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica uma cor de fundo específica em uma célula, linha ou intervalo usando
	 * valores RGB.
	 *
	 * @param red   Valor de vermelho (0-255).
	 * @param green Valor de verde (0-255).
	 * @param blue  Valor de azul (0-255).
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula corDeFundo(int red, int green, int blue) {
		backGroundColor.aplicarCorDeFundo(red, green, blue, rowIndex, columnIndex, startRowIndex, startColumnIndex,
				endRowIndex, endColumnIndex, isRange);
		return this;
	}

	/**
	 * Aplica uma cor de fundo específica em uma célula, linha ou intervalo usando
	 * um código hexadecimal.
	 *
	 * @param hexColor O código hexadecimal da cor de fundo (ex: "#FFFFFF").
	 * @return Instância atual de EstiloCelula para encadeamento de métodos.
	 */
	public EstiloCelula corDeFundo(String hexColor) {
		backGroundColor.aplicarCorDeFundo(hexColor, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
		return this;
	}

	// Getters para uso nas classes auxiliares (se necessário)

	private Workbook getWorkbook() {
		return workbook;
	}

	private Sheet getSheet() {
		return sheet;
	}

	private int getRowIndex() {
		return rowIndex;
	}

	private int getColumnIndex() {
		return columnIndex;
	}

	private int getStartRowIndex() {
		return startRowIndex;
	}

	private int getStartColumnIndex() {
		return startColumnIndex;
	}

	private int getEndRowIndex() {
		return endRowIndex;
	}

	private int getEndColumnIndex() {
		return endColumnIndex;
	}

	private boolean isRange() {
		return isRange;
	}

	private Map<String, org.apache.poi.ss.usermodel.CellStyle> getStyleCache() {
		return styleCache;
	}
}
