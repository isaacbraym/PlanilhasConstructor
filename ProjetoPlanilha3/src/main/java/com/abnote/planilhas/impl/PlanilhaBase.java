package com.abnote.planilhas.impl;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.interfaces.IManipulacaoDados;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.utils.LoggerUtil;
import com.abnote.planilhas.utils.ManipuladorPlanilha;
import com.abnote.planilhas.utils.PosicaoConverter;
import com.abnote.planilhas.utils.PositionManager;

public abstract class PlanilhaBase implements IPlanilha {
	private static final Logger logger = LoggerUtil.getLogger(PlanilhaBase.class);

	protected Workbook workbook;
	protected Sheet sheet;
	private final PositionManager positionManager = new PositionManager();
	private DataManipulator dataManipulator;
	private StyleManager styleManager;
	private String diretorioSaida = "C:\\opt\\tmp\\testePlanilhaSaidas";

	protected abstract void inicializarWorkbook();

	// Métodos de IPlanilhaBasica
	@Override
	public void criarPlanilha(String nomeSheet) {
		logger.info("Iniciando a criação da planilha: " + nomeSheet);
		try {
			inicializarWorkbook();
			sheet = workbook.createSheet(nomeSheet);
			positionManager.resetarPosicao();
			dataManipulator = new DataManipulator(workbook, sheet, positionManager);
			styleManager = new StyleManager(workbook, sheet, positionManager, dataManipulator);
			logger.info("Planilha '" + nomeSheet + "' criada com sucesso.");
		} catch (Exception e) {
			logger.severe("Erro ao criar a planilha '" + nomeSheet + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public void criarSheet(String nomeSheet) {
		try {
			if (workbook.getSheet(nomeSheet) != null) {
				String msg = "A aba '" + nomeSheet + "' já existe!";
				logger.warning(msg);
				throw new IllegalArgumentException(msg);
			}
			sheet = workbook.createSheet(nomeSheet);
			positionManager.resetarPosicao();
			dataManipulator = new DataManipulator(workbook, sheet, positionManager);
			styleManager = new StyleManager(workbook, sheet, positionManager, dataManipulator);
		} catch (Exception e) {
			logger.severe("Erro ao criar a aba '" + nomeSheet + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public void SELECIONAR_SHEET(String nomeSheet) {
		logger.fine("Atuando na Sheet: " + nomeSheet);
		try {
			if (workbook == null) {
				String msg = "Workbook ainda não foi inicializado!";
				logger.severe(msg);
				throw new IllegalStateException(msg);
			}

			sheet = workbook.getSheet(nomeSheet);

			if (sheet == null) {
				String msg = "A aba '" + nomeSheet + "' não foi encontrada.";
				logger.warning(msg);
				throw new IllegalArgumentException(msg);
			}

			positionManager.resetarPosicao();
			dataManipulator = new DataManipulator(workbook, sheet, positionManager);
			styleManager = new StyleManager(workbook, sheet, positionManager, dataManipulator);
		} catch (Exception e) {
			logger.severe("Erro ao selecionar a aba '" + nomeSheet + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public void salvar(String nomeArquivo) throws IOException {
		try (FileOutputStream arquivoSaida = new FileOutputStream(nomeArquivo)) {
			workbook.write(arquivoSaida);
			logger.info("Planilha salva com sucesso em: " + nomeArquivo);
		} catch (IOException e) {
			logger.severe("Erro ao salvar a planilha em '" + nomeArquivo + "': " + e.getMessage());
			throw e;
		}
	}

	@Override
	public void setDiretorioSaida(String diretorioSaida) {
		this.diretorioSaida = diretorioSaida;
	}

	@Override
	public String getDiretorioSaida() {
		return diretorioSaida;
	}

	@Override
	public Workbook obterWorkbook() {
		return workbook;
	}

	@Override
	public IPlanilha emTodaAPlanilha() {
		try {
			positionManager.emTodaAPlanilha();
			return this;
		} catch (Exception e) {
			logger.severe("Erro ao aplicar operações em toda a planilha: " + e.getMessage());
			throw e;
		}
	}

	@Override
	public int getNumeroDeLinhas(String coluna) {
		int colunaIndex = PosicaoConverter.converterColuna(coluna);
		int lastRowNum = sheet.getLastRowNum();
		int numRows = 0;
		for (int i = 0; i <= lastRowNum; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(colunaIndex);
				if (cell != null && cell.getCellTypeEnum() != CellType.BLANK) {
					numRows++;
				}
			}
		}
		return numRows;
	}

	@Override
	public int getNumeroDeColunasNaLinha(int linha) {
	    int linhaIndex = linha - 1;
	    Row row = sheet.getRow(linhaIndex);
	    if (row == null) {
	        return 0;
	    }
	    int numCols = 0;
	    short lastCellNum = row.getLastCellNum();
	    for (int i = 0; i < lastCellNum; i++) {
	        Cell cell = row.getCell(i);
	        if (cell != null && cell.getCellTypeEnum() != CellType.BLANK) {
	            numCols++;
	        }
	    }
	    return numCols;
	}


	@Override
	public IPlanilha ultimaLinha(String coluna) {
		try {
			dataManipulator.ultimaLinha(coluna);
			return this;
		} catch (Exception e) {
			logger.severe("Erro ao encontrar a última linha na coluna '" + coluna + "': " + e.getMessage());
			throw e;
		}
	}
	@Override
	public IManipulacaoDados naUltimaLinha(String coluna) {
	    try {
	        dataManipulator.naUltimaLinha(coluna);
	        return dataManipulator;
	    } catch (Exception e) {
	        logger.severe("Erro ao encontrar a última linha na coluna '" + coluna + "': " + e.getMessage());
	        throw e;
	    }
	}

	@Override
	public ManipuladorPlanilha manipularPlanilha() {
		return new ManipuladorPlanilha(sheet);
	}

	// Métodos de IManipulacaoDados delegados para DataManipulator

	@Override
	public IPlanilha naCelula(String posicao) {
		dataManipulator.naCelula(posicao);
		return this;
	}

	@Override
	public IPlanilha noIntervalo(String posicaoInicial, String posicaoFinal) {
		dataManipulator.noIntervalo(posicaoInicial, posicaoFinal);
		return this;
	}

	@Override
	public IPlanilha inserirDados(Object dados, String delimitador) {
		dataManipulator.inserirDados(dados, delimitador);
		return this;
	}

	@Override
	public IPlanilha inserirDados(String valor) {
		dataManipulator.inserirDados(valor);
		return this;
	}

	@Override
	public IPlanilha inserirDados(java.util.List<String> dados) {
		dataManipulator.inserirDados(dados);
		return this;
	}

	@Override
	public IPlanilha inserirDados(java.util.List<String> dados, String delimitador) {
		dataManipulator.inserirDados(dados, delimitador);
		return this;
	}

	@Override
	public IPlanilha inserirDadosArquivo(String caminhoArquivo, String delimitador) {
		try {
			dataManipulator.inserirDadosArquivo(caminhoArquivo, delimitador);
		} catch (Exception e) {
			logger.severe("Erro ao inserir dados do arquivo '" + caminhoArquivo + "': " + e.getMessage());
			throw e;
		}
		return this;
	}

	@Override
	public IPlanilha converterEmNumero(String posicaoInicial) {
		dataManipulator.converterEmNumero(posicaoInicial);
		return this;
	}

	@Override
	public IPlanilha converterEmContabil(String posicaoInicial) {
		dataManipulator.converterEmContabil(posicaoInicial);
		return this;
	}

	@Override
	public IPlanilha somarColuna(String posicaoInicial) {
		dataManipulator.somarColuna(posicaoInicial);
		return this;
	}

	@Override
	public IPlanilha somarColunaComTexto(String posicaoInicial, String texto) {
		dataManipulator.somarColunaComTexto(posicaoInicial, texto);
		return this;
	}

	@Override
	public IPlanilha multiplicarColunasComTexto(String coluna1, String coluna2, int linhaInicial, String texto,
			String colunaDestino) {
		dataManipulator.multiplicarColunasComTexto(coluna1, coluna2, linhaInicial, texto, colunaDestino);
		return this;
	}

	@Override
	public IPlanilha mesclarCelulas() {
		dataManipulator.mesclarCelulas();
		return this;
	}

	// Métodos de IEstilos delegados para StyleManager

	@Override
	public EstiloCelula aplicarEstilos() {
		return styleManager.aplicarEstilos();
	}

	@Override
	public EstiloCelula centralizarTudo() {
		return styleManager.centralizarTudo();
	}

	@Override
	public EstiloCelula redimensionarColunas() {
		return styleManager.redimensionarColunas();
	}

	@Override
	public EstiloCelula removerLinhasDeGrade() {
		return styleManager.removerLinhasDeGrade();
	}

	@Override
	public EstiloCelula aplicarEstilosEmCelula() {
		return styleManager.aplicarEstilosEmCelula();
	}

	@Override
	public EstiloCelula todasAsBordasEmTudo() {
		return styleManager.todasAsBordasEmTudo();
	}
}
