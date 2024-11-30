package com.abnote.planilhas.interfaces;

public interface IPlanilha extends IPlanilhaBasica, IEstilos, IManipulacaoDados {
	
	int getNumeroDeLinhas(String coluna);

	int getNumeroDeColunasNaLinha(int linha);
	
}
