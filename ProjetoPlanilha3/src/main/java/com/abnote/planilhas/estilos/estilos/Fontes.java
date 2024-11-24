package com.abnote.planilhas.estilos.estilos;

import java.awt.Color;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Fontes {
	private final Workbook workbook;
	private final Sheet sheet;
	private final Map<String, CellStyle> styleCache;

	public Fontes(Workbook workbook, Sheet sheet, Map<String, CellStyle> styleCache) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.styleCache = styleCache;
	}

	// Método genérico para aplicar atributos de fonte
	public void aplicarAtributosFonte(FontAttributes attributes, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {

		iterateCells(rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, isRange,
				(Cell cell) -> applyFontAttributesToCell(cell, attributes));
	}

	public void aplicarItalico(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setItalic(true);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarSublinhado(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setUnderline(Font.U_SINGLE);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarTachado(int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex, int endRowIndex,
			int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setStrikeout(true); // Agora funciona
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
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

	// Método para aplicar os atributos da fonte em uma célula
	private void applyFontAttributesToCell(Cell cell, FontAttributes attributes) {
		// Obter o estilo atual da célula
		CellStyle currentStyle = cell.getCellStyle();

		// Clonar o estilo atual
		CellStyle newStyle = styleCache.get("font_" + currentStyle.hashCode() + "_" + attributes.hashCode());
		if (newStyle == null) {
			newStyle = workbook.createCellStyle();
			newStyle.cloneStyleFrom(currentStyle);

			// Obter a fonte atual
			Font currentFont = workbook.getFontAt(currentStyle.getFontIndex());

			// Criar ou obter uma fonte com os novos atributos
			Font newFont = findOrCreateFont(currentFont, attributes);

			// Associar a nova fonte ao novo estilo
			newStyle.setFont(newFont);

			// Armazenar no cache
			styleCache.put("font_" + currentStyle.hashCode() + "_" + attributes.hashCode(), newStyle);
		}

		// Aplicar o novo estilo à célula
		cell.setCellStyle(newStyle);
	}

	private Font findOrCreateFont(Font currentFont, FontAttributes attributes) {
		// Verificar se já existe uma fonte com os atributos desejados
		for (short i = 0; i < workbook.getNumberOfFonts(); i++) {
			Font font = workbook.getFontAt(i);
			if (fontMatchesAttributes(font, currentFont, attributes)) {
				return font;
			}
		}

		// Se não encontrou, criar uma nova fonte clonando a atual e alterando os
		// atributos
		Font newFont = workbook.createFont();
		copyFontAttributes(newFont, currentFont, attributes);
		return newFont;
	}

	private boolean fontMatchesAttributes(Font font, Font baseFont, FontAttributes attributes) {
		if (attributes.getFontName() != null && !font.getFontName().equals(attributes.getFontName())) {
			return false;
		}
		if (attributes.getFontSize() != null && font.getFontHeightInPoints() != attributes.getFontSize()) {
			return false;
		}
		if (attributes.isBold() != null && font.getBold() != attributes.isBold()) {
			return false;
		}
		if (attributes.isItalic() != null && font.getItalic() != attributes.isItalic()) {
			return false;
		}
		if (attributes.getUnderline() != null && font.getUnderline() != attributes.getUnderline()) {
			return false;
		}
		if (attributes.isStrikeout() != null && font.getStrikeout() != attributes.isStrikeout()) {
			return false;
		}
		if (attributes.getColorRGB() != null && !fontsHaveSameRGBColor(font, attributes.getColorRGB())) {
			return false;
		}
		// Comparar outros atributos conforme necessário

		// Comparar atributos não alterados
		return fontsHaveSameBaseAttributes(font, baseFont, attributes);
	}

	private boolean fontsHaveSameBaseAttributes(Font font1, Font font2, FontAttributes attributes) {
		// Comparar atributos que não foram alterados
		if (attributes.getFontName() == null && !font1.getFontName().equals(font2.getFontName())) {
			return false;
		}
		if (attributes.getFontSize() == null && font1.getFontHeightInPoints() != font2.getFontHeightInPoints()) {
			return false;
		}
		if (attributes.isBold() == null && font1.getBold() != font2.getBold()) {
			return false;
		}
		if (attributes.isItalic() == null && font1.getItalic() != font2.getItalic()) {
			return false;
		}
		if (attributes.getUnderline() == null && font1.getUnderline() != font2.getUnderline()) {
			return false;
		}
		if (attributes.isStrikeout() == null && font1.getStrikeout() != font2.getStrikeout()) {
			return false;
		}
		if (attributes.getColorRGB() == null && !fontsHaveSameColor(font1, font2)) {
			return false;
		}
		// Comparar outros atributos conforme necessário
		return true;
	}

	private void copyFontAttributes(Font newFont, Font currentFont, FontAttributes attributes) {
		newFont.setFontName(attributes.getFontName() != null ? attributes.getFontName() : currentFont.getFontName());
		newFont.setFontHeightInPoints(
				attributes.getFontSize() != null ? attributes.getFontSize() : currentFont.getFontHeightInPoints());
		newFont.setBold(attributes.isBold() != null ? attributes.isBold() : currentFont.getBold());
		newFont.setItalic(attributes.isItalic() != null ? attributes.isItalic() : currentFont.getItalic());
		newFont.setUnderline(
				attributes.getUnderline() != null ? attributes.getUnderline() : currentFont.getUnderline());
		newFont.setStrikeout(attributes.isStrikeout() != null ? attributes.isStrikeout() : currentFont.getStrikeout());
		newFont.setCharSet(currentFont.getCharSet());
		newFont.setTypeOffset(currentFont.getTypeOffset());

		if (attributes.getColorRGB() != null) {
			if (newFont instanceof XSSFFont) {
				XSSFFont xssfFont = (XSSFFont) newFont;
				xssfFont.setColor(new XSSFColor(attributes.getColorRGB()));
			} else if (newFont instanceof HSSFFont) {
				// Para HSSF, mapeie para a cor indexada mais próxima
				short colorIndex = getNearestColorIndex(attributes.getColorRGB());
				newFont.setColor(colorIndex);
			}
		} else {
			newFont.setColor(currentFont.getColor());
			if (currentFont instanceof XSSFFont && newFont instanceof XSSFFont) {
				XSSFColor color = ((XSSFFont) currentFont).getXSSFColor();
				((XSSFFont) newFont).setColor(color);
			}
		}
	}

	private boolean fontsHaveSameColor(Font font1, Font font2) {
		if (font1 instanceof XSSFFont && font2 instanceof XSSFFont) {
			XSSFColor color1 = ((XSSFFont) font1).getXSSFColor();
			XSSFColor color2 = ((XSSFFont) font2).getXSSFColor();
			if (color1 == null && color2 == null) {
				return true;
			}
			return color1 != null && color1.equals(color2);
		} else {
			return font1.getColor() == font2.getColor();
		}
	}

	private boolean fontsHaveSameRGBColor(Font font, Color color) {
		if (font instanceof XSSFFont) {
			XSSFColor fontColor = ((XSSFFont) font).getXSSFColor();
			if (fontColor != null) {
				byte[] rgb = fontColor.getRGB();
				return rgb[0] == (byte) color.getRed() && rgb[1] == (byte) color.getGreen()
						&& rgb[2] == (byte) color.getBlue();
			}
		}
		return false;
	}

	private short getNearestColorIndex(Color color) {
		// Implementação para HSSF se necessário
		return IndexedColors.BLACK.getIndex(); // Retorne preto como padrão
	}

	// Métodos públicos para aplicar estilos específicos

	public void aplicarCorFonte(CorEnum corEnum, int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes()
				.setColorRGB(new Color(corEnum.getRed(), corEnum.getGreen(), corEnum.getBlue()));
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarCorFonte(String hexColor, int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		Color color = hexToColor(hexColor);
		FontAttributes attributes = new FontAttributes().setColorRGB(color);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarCorFonte(int red, int green, int blue, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		Color color = new Color(red, green, blue);
		FontAttributes attributes = new FontAttributes().setColorRGB(color);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarFonte(String fontName, int rowIndex, int columnIndex, int startRowIndex, int startColumnIndex,
			int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setFontName(fontName);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarFonte(FonteEnum fonteEnum, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		aplicarFonte(fonteEnum.getFontName(), rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	public void aplicarTamanhoFonte(int fontSize, int rowIndex, int columnIndex, int startRowIndex,
			int startColumnIndex, int endRowIndex, int endColumnIndex, boolean isRange) {
		FontAttributes attributes = new FontAttributes().setFontSize((short) fontSize);
		aplicarAtributosFonte(attributes, rowIndex, columnIndex, startRowIndex, startColumnIndex, endRowIndex,
				endColumnIndex, isRange);
	}

	// Método para converter código hexadecimal em Color
	private Color hexToColor(String hexColor) {
		if (hexColor == null || !hexColor.matches("^#([A-Fa-f0-9]{6})$")) {
			throw new IllegalArgumentException("Código hexadecimal de cor inválido: " + hexColor);
		}
		return Color.decode(hexColor);
	}

	// Criar métodos para aplicar outros estilos dealinhamentos,quebra de texto,
	// etc.
}
