package com.abnote.planilhas.utils;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Pattern;
import java.util.logging.Logger;

import com.abnote.planilhas.exceptions.ArquivoException;
import com.abnote.planilhas.exceptions.DadosInvalidosException;

import org.apache.poi.ss.usermodel.*;

public class InsersorDeDados {

	private static final Logger logger = LoggerUtil.getLogger(InsersorDeDados.class);

	private final Sheet sheet;
	private final PositionManager positionManager;
	private int ultimoIndiceDeLinhaInserido = -1;
	private int ultimoIndiceDeColunaInserido = -1;

	public InsersorDeDados(Sheet sheet, PositionManager positionManager) {
		this.sheet = sheet;
		this.positionManager = positionManager;
	}

	public void inserirDados(String valor) {
		definirPosicaoPadraoSeNecessario();

		Row linha = obterOuCriarLinha(positionManager.getPosicaoInicialLinha());
		Cell celula = linha.createCell(positionManager.getPosicaoInicialColuna());

		// Tenta converter para número, se falhar insere como string
		definirValorCelula(celula, valor);

		atualizarIndicesInseridos(positionManager.getPosicaoInicialLinha(), positionManager.getPosicaoInicialColuna());
	}

	public void inserirDados(Object dados, String delimitador) {
		if (dados instanceof List) {
			@SuppressWarnings("unchecked")
			List<String> lista = (List<String>) dados;
			inserirDados(lista);
		} else if (dados instanceof String) {
			String str = (String) dados;
			if (Files.exists(Paths.get(str))) {
				inserirDadosArquivo(str, delimitador);
			} else {
				List<String> lista = Arrays.asList(str.split(Pattern.quote(delimitador)));
				inserirDados(lista);
			}
		} else if (dados instanceof File) {
			inserirDadosArquivo(((File) dados).getPath(), delimitador);
		} else {
			throw new IllegalArgumentException("Tipo de dados não suportado: " + dados.getClass());
		}
	}

	public void inserirDados(List<String> dados) {
		definirPosicaoPadraoSeNecessario();

		if (positionManager.isIntervaloDefinida()) {
			inserirDadosEmIntervalo(dados);
		} else {
			inserirDadosEmLinha(dados);
		}

		positionManager.resetarPosicao();
	}

	public void inserirDados(List<String> dados, String delimitador) {
		inserirDados(dados);
	}

	public void inserirDadosArquivo(String caminhoArquivo, String delimitador) {
        definirPosicaoPadraoSeNecessario();
        
        if (caminhoArquivo == null || caminhoArquivo.trim().isEmpty()) {
            throw new ArquivoException(
                "Caminho do arquivo não pode ser nulo ou vazio",
                caminhoArquivo
            );
        }
        
        // [REMOVED] Validação do delimitador removida (pode ser vazio em casos válidos)

        try (BufferedReader br = new BufferedReader(new FileReader(caminhoArquivo))) {
            String linhaTexto;
            int linhaAtual = positionManager.getPosicaoInicialLinha();

            while ((linhaTexto = br.readLine()) != null) {
                String[] valores = linhaTexto.split(Pattern.quote(delimitador));
                inserirValoresEmLinha(linhaAtual, valores);

                linhaAtual++;
                if (positionManager.isIntervaloDefinida() && linhaAtual > positionManager.getPosicaoFinalLinha()) {
                    break;
                }
            }

            atualizarIndicesInseridos(linhaAtual - 1, positionManager.getPosicaoInicialColuna());
            positionManager.setPosicaoInicialLinha(linhaAtual);

        } catch (IOException e) {
            logger.severe("Erro ao ler o arquivo: " + e.getMessage());
            throw new ArquivoException(
                "Erro ao ler arquivo. Verifique se o arquivo existe e está acessível",
                caminhoArquivo,
                e
            );
        }

        positionManager.resetarPosicao();
    }

	/**
	 * Define o valor da célula, tentando converter para número quando possível.
	 * 
	 * @param celula Célula a receber o valor
	 * @param valor  String com o valor a ser inserido
	 */
	private void definirValorCelula(Cell celula, String valor) {
		if (valor == null || valor.trim().isEmpty()) {
			celula.setCellValue("");
			return;
		}

		String valorTrimmed = valor.trim();

		// Tenta converter para número
		try {
			double numeroDouble = Double.parseDouble(valorTrimmed);
			celula.setCellValue(numeroDouble);
		} catch (NumberFormatException e) {
			// Não é número, insere como string
			celula.setCellValue(valorTrimmed);
		}
	}

	// Métodos auxiliares privados

	private void definirPosicaoPadraoSeNecessario() {
		if (!positionManager.isPosicaoDefinida() && !positionManager.isIntervaloDefinida()) {
			positionManager.setPosicaoInicialColuna(0);
			positionManager.setPosicaoInicialLinha(0);
		}
	}

	private Row obterOuCriarLinha(int indiceLinha) {
		Row linha = sheet.getRow(indiceLinha);
		if (linha == null) {
			linha = sheet.createRow(indiceLinha);
		}
		return linha;
	}

	private void inserirDadosEmLinha(List<String> dados) {
		Row linha = obterOuCriarLinha(positionManager.getPosicaoInicialLinha());

		for (int i = 0; i < dados.size(); i++) {
			Cell celula = linha.createCell(positionManager.getPosicaoInicialColuna() + i);
			definirValorCelula(celula, dados.get(i)); // ✅ MUDANÇA AQUI
			ultimoIndiceDeColunaInserido = positionManager.getPosicaoInicialColuna() + i;
		}

		atualizarIndicesInseridos(positionManager.getPosicaoInicialLinha(), ultimoIndiceDeColunaInserido);
		positionManager.setPosicaoInicialLinha(positionManager.getPosicaoInicialLinha() + 1);
	}

	private void inserirDadosEmIntervalo(List<String> dados) {
		int linhaAtual = positionManager.getPosicaoInicialLinha();

		for (String dado : dados) {
			if (linhaAtual > positionManager.getPosicaoFinalLinha()) {
				break;
			}

			Row linha = obterOuCriarLinha(linhaAtual);

			for (int coluna = positionManager.getPosicaoInicialColuna(); coluna <= positionManager
					.getPosicaoFinalColuna(); coluna++) {
				Cell celula = linha.createCell(coluna);
				definirValorCelula(celula, dado); // ✅ MUDANÇA AQUI
			}

			linhaAtual++;
		}

		atualizarIndicesInseridos(linhaAtual - 1, positionManager.getPosicaoFinalColuna());
	}

	private void inserirValoresEmLinha(int indiceLinha, String[] valores) {
		Row linha = obterOuCriarLinha(indiceLinha);

		for (int i = 0; i < valores.length; i++) {
			int colunaAtual = positionManager.getPosicaoInicialColuna() + i;

			if (positionManager.isIntervaloDefinida() && colunaAtual > positionManager.getPosicaoFinalColuna()) {
				break;
			}

			Cell celula = linha.createCell(colunaAtual);
			definirValorCelula(celula, valores[i].trim()); // ✅ MUDANÇA AQUI
			ultimoIndiceDeColunaInserido = colunaAtual;
		}
	}

	private void atualizarIndicesInseridos(int linha, int coluna) {
		ultimoIndiceDeLinhaInserido = linha;
		ultimoIndiceDeColunaInserido = coluna;
	}

	// Getters para obter os últimos índices inseridos

	public int getUltimoIndiceDeLinhaInserido() {
		return ultimoIndiceDeLinhaInserido;
	}

	public int getUltimoIndiceDeColunaInserido() {
		return ultimoIndiceDeColunaInserido;
	}
}
