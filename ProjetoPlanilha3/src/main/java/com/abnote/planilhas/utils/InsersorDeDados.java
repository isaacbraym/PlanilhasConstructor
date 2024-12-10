package com.abnote.planilhas.utils;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.*;

public class InsersorDeDados {

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
		celula.setCellValue(valor);

		atualizarIndicesInseridos(positionManager.getPosicaoInicialLinha(), positionManager.getPosicaoInicialColuna());
//		positionManager.resetarPosicao();
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
			System.err.println("Erro ao ler o arquivo: " + e.getMessage());
		}

		positionManager.resetarPosicao();
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
			celula.setCellValue(dados.get(i));
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
				celula.setCellValue(dado);
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
			celula.setCellValue(valores[i].trim());
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
