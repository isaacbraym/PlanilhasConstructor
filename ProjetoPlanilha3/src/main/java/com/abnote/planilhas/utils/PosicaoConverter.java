package com.abnote.planilhas.utils;

public class PosicaoConverter {
	/**
	 * Converte uma posição no formato "A1" para índices de coluna e linha.
	 * 
	 * @param posicao Posição no formato "A1".
	 * @return Array com [coluna, linha].
	 */
	public static int[] converterPosicao(String posicao) {
		int coluna = 0;
		int linha = 0;
		for (int i = 0; i < posicao.length(); i++) {
			char ch = posicao.charAt(i);
			if (Character.isLetter(ch)) {
				coluna = coluna * 26 + (Character.toUpperCase(ch) - 'A' + 1);
			} else if (Character.isDigit(ch)) {
				linha = Integer.parseInt(posicao.substring(i)) - 1;
				break;
			}
		}
		return new int[] { coluna - 1, linha };
	}

	public static int converterColuna(String coluna) {
		coluna = coluna.toUpperCase();
		int length = coluna.length();
		int numero = 0;
		for (int i = 0; i < length; i++) {
			char c = coluna.charAt(i);
			numero = numero * 26 + (c - ('A' - 1));
		}
		return numero - 1; // Índice começa em 0
	}

	// Converte um índice numérico de coluna (0-based) em uma letra de coluna (por
	// exemplo, 0 -> "A")
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