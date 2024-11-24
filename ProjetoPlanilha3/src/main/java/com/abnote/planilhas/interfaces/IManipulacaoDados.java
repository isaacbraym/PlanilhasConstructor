package com.abnote.planilhas.interfaces;

import java.util.List;

import com.abnote.planilhas.estilos.EstiloCelula;

public interface IManipulacaoDados {

	EstiloCelula aplicarEstilos();

	EstiloCelula aplicarEstilosEmCelula();

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

}
