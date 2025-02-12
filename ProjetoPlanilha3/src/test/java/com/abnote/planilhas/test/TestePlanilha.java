package com.abnote.planilhas.test;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Logger;

import com.abnote.planilhas.estilos.estilos.CorEnum;
import com.abnote.planilhas.estilos.estilos.FonteEnum;
import com.abnote.planilhas.impl.PlanilhaXlsx;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.utils.LoggerUtil;

public class TestePlanilha {
    public static void main(String[] args) {
        final Logger logger = LoggerUtil.getLogger(TestePlanilha.class);

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
        planilha.naCelula("B2").inserirDados(header, ",");
        planilha.aplicarEstilos()
                .fonte(FonteEnum.TIMES_NEW_ROMAN)
                .corFonte(CorEnum.LARANJA)
                .fonteTamanho(14)
                .corDeFundo(CorEnum.VERMELHO_ESCURO)
                .aplicarNegrito();

        planilha.naCelula("B3").inserirDados(listaDeArquivos, ";");
        planilha.emTodaAPlanilha().aplicarEstilos().fonte("Segoe UI").fonteTamanho(14);
        planilha.aplicarEstilos().aplicarBordasEspessasComInternas("B2", "J2");
        planilha.naCelula("N11").inserirDados("TESTE");
        planilha.aplicarEstilosEmCelula().aplicarNegrito().corDeFundo(CorEnum.VERDE).corFonte(CorEnum.TURQUESA);

        planilha.converterEmContabil("J3").somarColunaComTexto("J3", "VALOR TOTAL DA SOMA");
        planilha.aplicarEstilos().aplicarItalico().aplicarNegrito();
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
        planilha.noIntervalo("G4", "H4").mesclarCelulas();
        planilha.aplicarEstilos().corDeFundo(CorEnum.VERMELHO_ESCURO).corFonte(CorEnum.BRANCO).aplicarNegrito();
        planilha.noIntervalo("C12", "C15").mesclarCelulas();
        planilha.aplicarEstilos().corDeFundo(CorEnum.AZUL_CELESTE).aplicarItalico().aplicarTachado();
        planilha.converterEmContabil("J3").multiplicarColunasComTexto("D", "I", 3, "Total multiplicação", "J")
                .aplicarEstilos().redimensionarColuna();
        planilha.ultimaLinha("I").aplicarEstilos().fonteTamanho(14).aplicarNegrito();
        // Exemplo de inserção de dados numéricos
        planilha.naUltimaLinha("E").inserir("TESTEE").aplicarEstilos().aplicarNegrito().fonteTamanho(14)
                .corDeFundo(CorEnum.LARANJA);
        planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();
        planilha.inserirFiltros();

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
        planilha.inserirFiltros();

        // PLANILHA SHEET3
        planilha.criarSheet(sheet3);
        planilha.SELECIONAR_SHEET(sheet3);
        planilha.inserirDados(listaDeArquivos2, "|");
        planilha.converterEmNumero("K2").somarColunaComTexto("K2", "Totais:");
        planilha.converterEmNumero("L2").somarColuna("L2").aplicarEstilos().aplicarNegrito();
        planilha.aplicarEstilos().removerLinhasDeGrade().centralizarERedimensionarTudo().aplicarTodasAsBordas();
        planilha.inserirFiltros();

        // Teste de logs (descomente se necessário)
        // logger.severe("Esta é uma mensagem SEVERE");
        // logger.warning("Esta é uma mensagem WARNING");
        // logger.info("Esta é uma mensagem INFO");
        // logger.config("Esta é uma mensagem CONFIG");
        // logger.fine("Esta é uma mensagem FINE");
        // logger.finer("Esta é uma mensagem FINER");
        // logger.finest("Esta é uma mensagem FINEST");

        try {
            planilha.salvar(caminhoArquivo);
        } catch (IOException e) {
            logger.severe("Erro ao salvar a planilha: " + e.getMessage());
        }
    }
}
