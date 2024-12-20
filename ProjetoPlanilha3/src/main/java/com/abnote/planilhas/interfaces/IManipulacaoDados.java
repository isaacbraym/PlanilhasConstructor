package com.abnote.planilhas.interfaces;

import java.util.List;

import com.abnote.planilhas.estilos.EstiloCelula;

public interface IManipulacaoDados {

	IManipulacaoDados naCelula(String posicao);

	IManipulacaoDados noIntervalo(String posicaoInicial, String posicaoFinal);

	IManipulacaoDados inserirDados(Object dados, String delimitador);

	IManipulacaoDados inserirDados(String valor);

	IManipulacaoDados inserirDados(List<String> dados);

	IManipulacaoDados inserirDados(List<String> dados, String delimitador);

	IManipulacaoDados inserirDadosArquivo(String caminhoArquivo, String delimitador);

	IManipulacaoDados converterEmNumero(String posicaoInicial);

	IManipulacaoDados converterEmContabil(String coluna);

	IManipulacaoDados somarColuna(String posicaoInicial);

	IManipulacaoDados somarColunaComTexto(String posicaoInicial, String texto);

	IManipulacaoDados multiplicarColunasComTexto(String coluna1, String coluna2, int linhaInicial, String texto,
			String colunaDestino);

	IManipulacaoDados mesclarCelulas();
	
    IManipulacaoDados inserir(String valor);
    IManipulacaoDados inserir(int valor);
    IManipulacaoDados inserir(double valor);

	EstiloCelula aplicarEstilos();

	IManipulacaoDados naUltimaLinha(String coluna);
}
