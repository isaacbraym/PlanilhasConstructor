package com.abnote.planilhas.utils;

public class PositionManager {
	private int posicaoInicialColuna = 0;
	private int posicaoInicialLinha = 0;
	private int posicaoFinalColuna = 0;
	private int posicaoFinalLinha = 0;
	private boolean intervaloDefinida = false;
	private boolean posicaoDefinida = false;
	private boolean todaPlanilhaDefinida = false;

	public void naCelula(String posicao) {
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

		this.posicaoInicialColuna = coluna - 1;
		this.posicaoInicialLinha = linha;
		this.posicaoDefinida = true;
	}

	public void noIntervalo(String posicaoInicial, String posicaoFinal) {
		int[] inicio = PosicaoConverter.converterPosicao(posicaoInicial);
		int[] fim = PosicaoConverter.converterPosicao(posicaoFinal);

		this.posicaoInicialColuna = inicio[0];
		this.posicaoInicialLinha = inicio[1];
		this.posicaoFinalColuna = fim[0];
		this.posicaoFinalLinha = fim[1];

		this.intervaloDefinida = true;
	}

	public void emTodaAPlanilha() {
		this.todaPlanilhaDefinida = true;
	}

	public boolean isTodaPlanilhaDefinida() {
		return todaPlanilhaDefinida;
	}

	public void resetarPosicao() {
		this.posicaoInicialColuna = 0;
		this.posicaoInicialLinha = 0;
		this.posicaoFinalColuna = 0;
		this.posicaoFinalLinha = 0;
		this.posicaoDefinida = false;
		this.intervaloDefinida = false;
		this.todaPlanilhaDefinida = false;
	}

	// Getters e setters
	public int getPosicaoInicialColuna() {
		return posicaoInicialColuna;
	}

	public void setPosicaoInicialColuna(int posicaoInicialColuna) {
		this.posicaoInicialColuna = posicaoInicialColuna;
	}

	public int getPosicaoInicialLinha() {
		return posicaoInicialLinha;
	}

	public void setPosicaoInicialLinha(int posicaoInicialLinha) {
		this.posicaoInicialLinha = posicaoInicialLinha;
	}

	public int getPosicaoFinalColuna() {
		return posicaoFinalColuna;
	}

	public int getPosicaoFinalLinha() {
		return posicaoFinalLinha;
	}

	public boolean isIntervaloDefinida() {
		return intervaloDefinida;
	}

	public boolean isPosicaoDefinida() {
		return posicaoDefinida;
	}
}
