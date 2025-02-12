package com.abnote.planilhas.utils;

/**
 * Classe utilitária para conversão entre a notação de posição (por exemplo,
 * "B2") e índices numéricos (base 0).
 */
public class PosicaoConverter {

	/**
	 * Converte uma posição no formato "B2" em um array de inteiros onde o primeiro
	 * elemento é o índice da coluna (0-based) e o segundo é o índice da linha
	 * (0-based).
	 *
	 * @param posicao A posição no formato alfanumérico.
	 * @return Array de inteiros: [coluna, linha].
	 */
	public static int[] converterPosicao(String posicao) {
		String colunaParte = posicao.replaceAll("\\d", "");
		String linhaParte = posicao.replaceAll("\\D", "");
		int coluna = converterColuna(colunaParte);
		int linha = Integer.parseInt(linhaParte) - 1;
		return new int[] { coluna, linha };
	}

	/**
	 * Converte uma letra ou conjunto de letras representando a coluna em um índice
	 * numérico (0-based).
	 *
	 * @param coluna A(s) letra(s) que representam a coluna (por exemplo, "A", "B",
	 *               "AA").
	 * @return Índice numérico da coluna.
	 */
	public static int converterColuna(String coluna) {
		coluna = coluna.toUpperCase();
		int length = coluna.length();
		int numero = 0;
		for (int i = 0; i < length; i++) {
			char c = coluna.charAt(i);
			numero = numero * 26 + (c - ('A' - 1));
		}
		return numero - 1; // Ajusta para índice 0-based
	}

	/**
	 * Converte um índice numérico de coluna (0-based) em sua representação
	 * alfabética (por exemplo, 0 -> "A").
	 *
	 * @param index Índice numérico da coluna.
	 * @return Representação alfabética da coluna.
	 */
	public static String converterIndice(int index) {
		StringBuilder result = new StringBuilder();
		index += 1; // Ajusta para 1-based
		while (index > 0) {
			int remainder = (index - 1) % 26;
			result.insert(0, (char) (remainder + 'A'));
			index = (index - 1) / 26;
		}
		return result.toString();
	}
}
