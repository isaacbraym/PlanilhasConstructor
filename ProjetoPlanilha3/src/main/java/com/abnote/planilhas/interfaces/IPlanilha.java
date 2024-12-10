package com.abnote.planilhas.interfaces;

public interface IPlanilha extends IPlanilhaBasica, IEstilos, IManipulacaoDados {
	
	IPlanilha inserirFiltros();
	
	int getNumeroDeLinhas(String coluna);

	int getNumeroDeColunasNaLinha(int linha);
	
}
