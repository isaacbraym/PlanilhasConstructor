package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.*;
import java.util.*;

public class ManipuladorPlanilha {
	private Sheet sheet;
	private final LogsDeModificadores logs;
	private Map<Integer, ManipuladorPlanilhaHelper.CellData> colunaTemporaria = new HashMap<>();
	private int columnOffset;
	private ManipuladorPlanilhaHelper helper;

	/**
	 * Construtor que determina o columnOffset automaticamente.
	 *
	 * @param sheet A planilha a ser manipulada.
	 */
	public ManipuladorPlanilha(Sheet sheet) {
		this(sheet, ManipuladorPlanilhaHelper.determinarColunaInicial(sheet));
	}

	/**
	 * Construtor que permite configurar o columnOffset manualmente.
	 *
	 * @param sheet        A planilha a ser manipulada.
	 * @param columnOffset O deslocamento a ser aplicado nas operações de coluna.
	 */
	public ManipuladorPlanilha(Sheet sheet, int columnOffset) {
		this.sheet = sheet;
		this.columnOffset = columnOffset;
		this.logs = new LogsDeModificadores();
		this.helper = new ManipuladorPlanilhaHelper(sheet, columnOffset);
	}

	/**
	 * Setter para ajustar o offset se necessário.
	 *
	 * @param columnOffset O novo deslocamento a ser aplicado.
	 */
	private void setColumnOffset(int columnOffset) {
		this.columnOffset = columnOffset;
		// Atualizar o helper com o novo offset
		this.helper = new ManipuladorPlanilhaHelper(sheet, columnOffset);
	}

	/**
	 * Método para mover uma coluna de uma posição para outra.
	 *
	 * @param moverAColuna Nome da coluna a ser movida (ex: "A", "B", etc.).
	 * @param paraAPosicao Nome da coluna para a qual será movida (ex: "C", "D",
	 *                     etc.).
	 * @return Instância atual de ManipuladorPlanilha para encadeamento de métodos.
	 */
	public ManipuladorPlanilha moverColuna(String moverAColuna, String paraAPosicao) {
		int colunaOrigem = PosicaoConverter.converterColuna(moverAColuna) - columnOffset;
		int colunaDestino = PosicaoConverter.converterColuna(paraAPosicao) - columnOffset;

		if (colunaOrigem == colunaDestino) {
			return this;
		}

		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos();
		String headerOrigem = headerMap.get(colunaOrigem);
		LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerOrigem,
				PosicaoConverter.converterIndice(colunaOrigem + columnOffset),
				PosicaoConverter.converterIndice(colunaDestino + columnOffset));
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Deslocamento de colunas",
				mainMovement);

		colunaTemporaria = helper.copiarColuna(colunaOrigem);

		if (colunaOrigem < colunaDestino) {
			helper.deslocarColunasParaEsquerda(colunaOrigem + 1, colunaDestino);
			helper.registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
		} else {
			helper.deslocarColunasParaDireita(colunaDestino, colunaOrigem - 1);
			helper.registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
		}

		helper.colarColunaTemporaria(colunaDestino, colunaTemporaria);
		colunaTemporaria.clear();
		logs.adicionarLog(actionLog);

		return this;
	}

	/**
	 * Método para remover uma coluna específica.
	 *
	 * @param coluna Nome da coluna a ser removida (ex: "A", "B", etc.).
	 * @return Instância atual de ManipuladorPlanilha para encadeamento de métodos.
	 */
	public ManipuladorPlanilha removerColuna(String coluna) {
		int colIndex = PosicaoConverter.converterColuna(coluna) - columnOffset;
		int lastColumn = helper.obterNumeroUltimaColuna();

		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos();
		String headerName = headerMap.get(colIndex);
		LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerName,
				PosicaoConverter.converterIndice(colIndex + columnOffset), null);
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Remoção de coluna", mainMovement);

		helper.removerCelulasDaColuna(colIndex);

		if (colIndex < lastColumn) {
			helper.deslocarColunasParaEsquerda(colIndex + 1, lastColumn); // Atualizado
			helper.registrarColunasDeslocadasRemocao(colIndex, lastColumn, headerMap, actionLog);
		}

		logs.adicionarLog(actionLog);
		return this;
	}

	/**
	 * Método para limpar os dados de uma coluna sem remover ou deslocar a coluna.
	 *
	 * @param coluna Nome da coluna a ser limpa (ex: "A", "B", etc.).
	 * @return Instância atual de ManipuladorPlanilha para encadeamento de métodos.
	 */
	public ManipuladorPlanilha limparColuna(String coluna) {
		int colIndex = PosicaoConverter.converterColuna(coluna) - columnOffset;

		// Utilizando o método helper para limpar a coluna
		helper.limparColuna(colIndex);

		// Registrar a ação no log
		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos(); // Atualizado
		String headerName = headerMap.get(colIndex);
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Limpeza de coluna",
				new LogsDeModificadores.ColumnMovement(headerName,
						PosicaoConverter.converterIndice(colIndex + columnOffset), null));
		logs.adicionarLog(actionLog);

		return this;
	}

	/**
	 * Método para inserir uma coluna vazia entre duas colunas especificadas.
	 *
	 * @param colunaEsquerda Nome da coluna à esquerda (ex: "A", "B", etc.).
	 * @param colunaDireita  Nome da coluna à direita (ex: "C", "D", etc.).
	 * @return Instância atual de ManipuladorPlanilha para encadeamento de métodos.
	 */
	public ManipuladorPlanilha inserirColunaVaziaEntre(String colunaEsquerda, String colunaDireita) {
		int colEsquerdaIndex = PosicaoConverter.converterColuna(colunaEsquerda) - columnOffset;
		int colDireitaIndex = PosicaoConverter.converterColuna(colunaDireita) - columnOffset;

		helper.validarAdjacencia(colEsquerdaIndex, colDireitaIndex, colunaEsquerda, colunaDireita);

		Map<Integer, String> headerMap = helper.obterMapaDeCabecalhos();
		int posicaoInsercao = colDireitaIndex;
		int lastColumn = helper.obterNumeroUltimaColuna();

		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Inserção de coluna vazia",
				new LogsDeModificadores.ColumnMovement(null, colunaEsquerda, colunaDireita));

		if (posicaoInsercao <= lastColumn) {
			helper.deslocarColunasParaDireita(posicaoInsercao, lastColumn); // Atualizado
			helper.registrarColunasDeslocadasInsercao(posicaoInsercao, lastColumn, headerMap, actionLog); // Atualizado
			logs.adicionarLog(actionLog);
		}

		helper.definirLarguraNovaColuna(posicaoInsercao);
		return this;
	}

	/**
	 * Método para exibir o log de rastreio.
	 */
	public void logAlteracoes() {
		logs.exibirLogs();
	}
}
