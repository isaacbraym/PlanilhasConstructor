package com.abnote.planilhas.test;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.estilos.estilos.FonteEnum;
import com.abnote.planilhas.impl.PlanilhaXlsx;
import com.abnote.planilhas.interfaces.IPlanilha;

public class TestePlanilha {
	public static void main(String[] args) {

		IPlanilha planilha = new PlanilhaXlsx();
		String sheet1 = "Dados Brasileiros";
		String sheet2 = "TesteAba2";
		String sheet3 = "TesteAba3";

		String diretorioSaida = "C:\\opt\\tmp\\testePlanilhaSaidas";
		planilha.setDiretorioSaida(diretorioSaida);

		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
		String nomeArquivoPlanilha = "\\planilhaTeste_" + timestamp + ".xlsx";
		String caminhoArquivo = planilha.getDiretorioSaida() + nomeArquivoPlanilha;
		String listaDeArquivos = "C:\\opt\\lista_dados_brasileiros.txt";
		String listaDeArquivos2 = "C:\\opt\\listaDados2.csv";

		String header = "NOME,CPF/CNPJ,ENDERECO,NUMERO,COMPLEMENTO,CEP,CIDADE,ESTADO,VALOR";

		// PLANILHA SHEET1
		planilha.criarPlanilha(sheet1);
		planilha.SELECIONAR_SHEET(sheet1);
		planilha.naCelula("B2").inserirDados(header, ",").aplicarEstilos().fonte(FonteEnum.TIMES_NEW_ROMAN)
				.corFonte(CorEnum.LARANJA).fonteTamanho(14).corDeFundo(CorEnum.VERMELHO_ESCURO).aplicarNegrito();
		planilha.naCelula("B3").inserirDados(listaDeArquivos, ";");
		planilha.emTodaAPlanilha().aplicarEstilos().fonte("Segoe UI").fonteTamanho(14);
		planilha.aplicarEstilos().aplicarBordasEspessasComInternas("B2", "J2");
		planilha.naCelula("N11").inserirDados("TESTE").aplicarEstilosEmCelula().aplicarNegrito()
				.corDeFundo(CorEnum.VERDE).corFonte(CorEnum.TURQUESA);
		planilha.converterEmContabil("J3").somarColunaComTexto("J3", "VALOR TOTAL DA SOMA").aplicarEstilos()
				.aplicarItalico().aplicarNegrito();
		planilha.aplicarEstilos().aplicarBordasEntre("L2", "L200");
		planilha.noIntervalo("C4", "C17").aplicarEstilos().aplicarNegrito();
		planilha.noIntervalo("C5", "G5").aplicarEstilos().fonte("Calibri").fonteTamanho(18).aplicarNegrito();
		planilha.noIntervalo("C4", "F4").aplicarEstilos().fonte(FonteEnum.EBRIMA).fonteTamanho(21).aplicarNegrito();
		planilha.manipularPlanilha().moverColuna("C", "F").logAlteracoes();
		planilha.manipularPlanilha().removerColuna("I").logAlteracoes();
		planilha.noIntervalo("G10", "G20").aplicarEstilos().corDeFundo(CorEnum.ROXO).corFonte(CorEnum.BRANCO);
		planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();
		planilha.noIntervalo("G15", "G20").aplicarEstilos().aplicarItalico().aplicarNegrito().alinharADireita();
		planilha.noIntervalo("G18", "G22").aplicarEstilos().aplicarNegrito().aplicarSublinhado().alinharAEsquerda();
		planilha.noIntervalo("F20", "H20").aplicarEstilos().aplicarTachado().aplicarNegrito();
		planilha.noIntervalo("G4", "H4").mesclarCelulas().aplicarEstilos().corDeFundo(CorEnum.VERMELHO_ESCURO).corFonte(CorEnum.BRANCO).aplicarNegrito();
		planilha.noIntervalo("C12", "C15").mesclarCelulas().aplicarEstilos().corDeFundo(CorEnum.AZUL_CELESTE).aplicarItalico().aplicarTachado();
//		planilha.manipularPlanilha().inserirColunaVaziaEntre("D", "E").logAlteracoes();
//		planilha.manipularPlanilha().limparColuna("D").logAlteracoes();
//		planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();

		// PLANILHA SHEET2
		planilha.criarSheet(sheet2);
		planilha.SELECIONAR_SHEET(sheet2);
		planilha.naCelula("C3").inserirDados(listaDeArquivos, ";");
		planilha.ultimaLinha("J").aplicarEstilos().aplicarNegrito().fonte("Arial").fonteTamanho(14);
		planilha.noIntervalo("C4", "C17").aplicarEstilos().fonte("Another Danger - Demo").fonteTamanho(12)
				.corDeFundo(CorEnum.AMARELO).aplicarNegrito();
		planilha.noIntervalo("C4", "F4").aplicarEstilos().fonte(FonteEnum.VERDANA).fonteTamanho(21)
				.corFonte(CorEnum.BRANCO).corDeFundo("#9400d3").aplicarNegrito();
		planilha.noIntervalo("D11", "G11").aplicarEstilos().corFonte(CorEnum.BEGE).corDeFundo(90, 50, 128)
				.aplicarNegrito();
		planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();

		// PLANILHA SHEET3
		planilha.criarSheet(sheet3);
		planilha.SELECIONAR_SHEET(sheet3);
		planilha.inserirDados(listaDeArquivos2, "|");
		planilha.converterEmNumero("K2").somarColunaComTexto("K2", "Totais:");
		planilha.converterEmNumero("L2").somarColuna("L2").aplicarEstilos().aplicarNegrito();
		planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();

		try {
			planilha.salvar(caminhoArquivo);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
