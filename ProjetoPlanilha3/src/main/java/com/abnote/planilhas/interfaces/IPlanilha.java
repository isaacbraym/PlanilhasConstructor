package com.abnote.planilhas.interfaces;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.utils.ManipuladorPlanilha;

public interface IPlanilha {
	void criarPlanilha(String nomeSheet);

	String getDiretorioSaida();

	Workbook obterWorkbook();

	// Métodos para evitar casts
	IPlanilha naCelula(String posicao);

	IPlanilha noIntervalo(String posicaoInicial, String posicaoFinal);

	IPlanilha inserirDados(Object dados, String delimitador);

	IPlanilha inserirDados(String valor);

	IPlanilha inserirDados(List<String> dados);

	IPlanilha inserirDados(List<String> dados, String delimitador);

	IPlanilha inserirDadosArquivo(String caminhoArquivo, String delimitador);

	IPlanilha converterEmNumero(String posicaoInicial);

	IPlanilha converterEmContabil(String coluna);

	IPlanilha somarColuna(String posicaoInicial);

	IPlanilha somarColunaComTexto(String posicaoInicial, String texto);

//	IPlanilha inserirEmLinha(String posicaoInicial, String posicaoFinal);

	IPlanilha ultimaLinha(String coluna);

	IPlanilha todasAsBordasEmTudo();

	EstiloCelula aplicarEstilos();

	EstiloCelula centralizarTudo();

	EstiloCelula redimensionarColunas();

	EstiloCelula removerLinhasDeGrade();

	EstiloCelula aplicarEstilosEmCelula();

	void criarSheet(String nomeSheet);

	void setDiretorioSaida(String diretorioSaida);

	void SELECIONAR_SHEET(String nomeSheet);

	void salvar(String nomeArquivo) throws IOException;

	IPlanilha emTodaAPlanilha();
	ManipuladorPlanilha manipularPlanilha();
//	void selecionarSheet(String nomeSheet);

}
