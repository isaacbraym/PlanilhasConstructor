package com.abnote.planilhas.utils;

import org.apache.poi.ss.usermodel.*;
import java.util.*;

public class ManipuladorPlanilha {
	private Sheet sheet;
	private final LogsDeModificadores logs;
	private Map<Integer, ManipuladorPlanilhaHelper.CellData> colunaTemporaria = new HashMap<>();
	private int columnOffset;
	private ManipuladorPlanilhaHelper helper;

	// Construtor que determina o columnOffset automaticamente
	public ManipuladorPlanilha(Sheet sheet) {
		this(sheet, ManipuladorPlanilhaHelper.determinarColunaInicial(sheet));
	}

	// Construtor que permite configurar o columnOffset manualmente
	public ManipuladorPlanilha(Sheet sheet, int columnOffset) {
		this.sheet = sheet;
		this.columnOffset = columnOffset;
		this.logs = new LogsDeModificadores();
		this.helper = new ManipuladorPlanilhaHelper(sheet, columnOffset);
	}

	// Setter para ajustar o offset se necessário
	public void setColumnOffset(int columnOffset) {
		this.columnOffset = columnOffset;
		// Atualizar o helper com o novo offset
		this.helper = new ManipuladorPlanilhaHelper(sheet, columnOffset);
	}

	// Método para mover uma coluna
	public ManipuladorPlanilha moverColuna(String moverAColuna, String paraAPosicao) {
		int colunaOrigem = PosicaoConverter.converterColuna(moverAColuna) - columnOffset;
		int colunaDestino = PosicaoConverter.converterColuna(paraAPosicao) - columnOffset;

		if (colunaOrigem == colunaDestino) {
			return this;
		}

		Map<Integer, String> headerMap = helper.getHeaderMap();
		String headerOrigem = headerMap.get(colunaOrigem);
		LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerOrigem,
				PosicaoConverter.converterIndice(colunaOrigem + columnOffset),
				PosicaoConverter.converterIndice(colunaDestino + columnOffset));
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Deslocamento de colunas",
				mainMovement);

		colunaTemporaria = helper.copiarColuna(colunaOrigem);

		if (colunaOrigem < colunaDestino) {
			helper.shiftColumnsLeft(colunaOrigem + 1, colunaDestino);
			helper.registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
		} else {
			helper.shiftColumnsRight(colunaDestino, colunaOrigem - 1);
			helper.registrarColunasDeslocadas(colunaOrigem, colunaDestino, headerMap, actionLog);
		}

		helper.colarColunaTemporaria(colunaDestino, colunaTemporaria);
		colunaTemporaria.clear();
		logs.adicionarLog(actionLog);

		return this;
	}

	// Método para remover uma coluna
	public ManipuladorPlanilha removerColuna(String coluna) {
		int colIndex = PosicaoConverter.converterColuna(coluna) - columnOffset;
		int lastColumn = helper.getLastColumnNum();

		Map<Integer, String> headerMap = helper.getHeaderMap();
		String headerName = headerMap.get(colIndex);
		LogsDeModificadores.ColumnMovement mainMovement = new LogsDeModificadores.ColumnMovement(headerName,
				PosicaoConverter.converterIndice(colIndex + columnOffset), null);
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Remoção de coluna", mainMovement);

		helper.removerCelulasDaColuna(colIndex);
		if (colIndex < lastColumn) {
			helper.shiftColumnsLeft(colIndex + 1, lastColumn);
			helper.registrarColunasDeslocadasRemocao(colIndex, lastColumn, headerMap, actionLog);
		}

		logs.adicionarLog(actionLog);
		return this;
	}

	/**
	 * Método para limpar os dados de uma coluna sem remover ou deslocar a coluna.
	 *
	 * @param coluna Nome da coluna a ser limpa (ex: "A", "B", etc.)
	 * @return Instância atual de ManipuladorPlanilha para encadeamento de métodos
	 */
	public ManipuladorPlanilha limparColuna(String coluna) {
		int colIndex = PosicaoConverter.converterColuna(coluna) - columnOffset;
		int lastRowNum = sheet.getLastRowNum();

		for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row != null) {
				Cell cell = row.getCell(colIndex + columnOffset);
				if (cell != null) {
					cell.setCellType(CellType.BLANK);
				}
			}
		}

		// Registrar a ação no log
		Map<Integer, String> headerMap = helper.getHeaderMap();
		String headerName = headerMap.get(colIndex);
		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Limpeza de coluna",
				new LogsDeModificadores.ColumnMovement(headerName,
						PosicaoConverter.converterIndice(colIndex + columnOffset), null));
		logs.adicionarLog(actionLog);

		return this;
	}

	// Método para inserir uma coluna vazia entre duas colunas especificadas
	public ManipuladorPlanilha inserirColunaVaziaEntre(String colunaEsquerda, String colunaDireita) {
		int colEsquerdaIndex = PosicaoConverter.converterColuna(colunaEsquerda) - columnOffset;
		int colDireitaIndex = PosicaoConverter.converterColuna(colunaDireita) - columnOffset;

		helper.validarAdjacencia(colEsquerdaIndex, colDireitaIndex, colunaEsquerda, colunaDireita);

		Map<Integer, String> headerMap = helper.getHeaderMap();
		int posicaoInsercao = colDireitaIndex;
		int lastColumn = helper.getLastColumnNum();

		LogsDeModificadores.ActionLog actionLog = new LogsDeModificadores.ActionLog("Inserção de coluna vazia",
				new LogsDeModificadores.ColumnMovement(null, colunaEsquerda, colunaDireita));

		if (posicaoInsercao <= lastColumn) {
			helper.shiftColumnsRight(posicaoInsercao, lastColumn);
			helper.registrarColunasDeslocadasInsercao(posicaoInsercao, lastColumn, headerMap, actionLog);
			logs.adicionarLog(actionLog);
		}

		helper.definirLarguraDaNovaColuna(posicaoInsercao);
		return this;
	}

	// Método para exibir o log de rastreio
	public void logAlteracoes() {
		logs.exibirLogs();
	}
}
