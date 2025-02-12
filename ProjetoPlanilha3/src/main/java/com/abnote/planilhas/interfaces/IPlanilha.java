package com.abnote.planilhas.interfaces;

/**
 * Interface que agrega as funcionalidades básicas, de estilos e de manipulação
 * de dados da planilha.
 */
public interface IPlanilha extends IPlanilhaBasica, IEstilos, IManipulacaoDados {

	/**
	 * Insere filtros na planilha, definindo o cabeçalho automaticamente.
	 *
	 * @return A instância atual para encadeamento.
	 */
	IPlanilha inserirFiltros();

	/**
	 * Retorna o número de linhas preenchidas na coluna informada.
	 *
	 * @param coluna A coluna (ex.: "B").
	 * @return O número de linhas.
	 */
	int getNumeroDeLinhas(String coluna);

	/**
	 * Retorna o número de colunas preenchidas na linha informada.
	 *
	 * @param linha A linha (1-based).
	 * @return O número de colunas.
	 */
	int getNumeroDeColunasNaLinha(int linha);
}
