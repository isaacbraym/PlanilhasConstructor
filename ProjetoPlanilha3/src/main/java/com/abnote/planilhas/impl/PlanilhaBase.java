package com.abnote.planilhas.impl;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.*;

import com.abnote.planilhas.calculos.Calculos;
import com.abnote.planilhas.calculos.Conversores;
import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.interfaces.IPlanilha;
import com.abnote.planilhas.interfaces.IPlanilhaBasica;
import com.abnote.planilhas.interfaces.IManipulacaoDados;
import com.abnote.planilhas.utils.InsersorDeDados;
import com.abnote.planilhas.utils.ManipuladorPlanilha;
import com.abnote.planilhas.utils.PosicaoConverter;
import com.abnote.planilhas.utils.PositionManager;

public abstract class PlanilhaBase implements IPlanilha {
    protected Workbook workbook;
    protected Sheet sheet;
    private final PositionManager positionManager = new PositionManager();
    private InsersorDeDados insersorDeDados;
    private String diretorioSaida = "C:\\opt\\tmp\\testePlanilhaSaidas";

    private int ultimoIndiceDeLinhaInserido = -1;
    private int ultimoIndiceDeColunaInserido = -1;

    protected abstract void inicializarWorkbook();

    // Métodos de IPlanilhaBasica

    @Override
    public void criarPlanilha(String nomeSheet) {
        inicializarWorkbook();
        sheet = workbook.createSheet(nomeSheet);
        insersorDeDados = new InsersorDeDados(sheet, positionManager);
        positionManager.resetarPosicao();
    }

    @Override
    public void criarSheet(String nomeSheet) {
        if (workbook.getSheet(nomeSheet) != null) {
            throw new IllegalArgumentException("A aba '" + nomeSheet + "' já existe!");
        }
        sheet = workbook.createSheet(nomeSheet);
        insersorDeDados = new InsersorDeDados(sheet, positionManager);
        positionManager.resetarPosicao();
    }

    @Override
    public void SELECIONAR_SHEET(String nomeSheet) {
        if (workbook == null) {
            throw new IllegalStateException("Workbook ainda não foi inicializado!");
        }

        sheet = workbook.getSheet(nomeSheet);

        if (sheet == null) {
            throw new IllegalArgumentException("A aba '" + nomeSheet + "' não foi encontrada.");
        }

        insersorDeDados = new InsersorDeDados(sheet, positionManager);
        positionManager.resetarPosicao();
    }

    @Override
    public void salvar(String nomeArquivo) throws IOException {
        try (FileOutputStream arquivoSaida = new FileOutputStream(nomeArquivo)) {
            workbook.write(arquivoSaida);
            System.out.println("Planilha criada com sucesso em: " + nomeArquivo);
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
    public IPlanilhaBasica emTodaAPlanilha() {
        positionManager.emTodaAPlanilha();
        return this;
    }

    @Override
    public IPlanilhaBasica ultimaLinha(String coluna) {
        int[] posicao = PosicaoConverter.converterPosicao(coluna + "1");
        int colunaIndex = posicao[0];

        int ultimaLinha = -1;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(colunaIndex);
                if (cell != null && cell.getCellTypeEnum() != CellType.BLANK) {
                    ultimaLinha = i;
                }
            }
        }

        ultimoIndiceDeLinhaInserido = (ultimaLinha >= 0) ? ultimaLinha : sheet.getLastRowNum();
        return this;
    }

    @Override
    public ManipuladorPlanilha manipularPlanilha() {
        return new ManipuladorPlanilha(sheet);
    }

    // Métodos de IManipulacaoDados

    @Override
    public IManipulacaoDados naCelula(String posicao) {
        positionManager.naCelula(posicao);
        return this;
    }

    @Override
    public IManipulacaoDados noIntervalo(String posicaoInicial, String posicaoFinal) {
        positionManager.noIntervalo(posicaoInicial, posicaoFinal);
        return this;
    }

    @Override
    public IManipulacaoDados inserirDados(Object dados, String delimitador) {
        insersorDeDados.inserirDados(dados, delimitador);
        updateLastInsertedIndices();
        return this;
    }

    @Override
    public IManipulacaoDados inserirDados(String valor) {
        insersorDeDados.inserirDados(valor);
        updateLastInsertedIndices();
        return this;
    }

    @Override
    public IManipulacaoDados inserirDados(List<String> dados) {
        insersorDeDados.inserirDados(dados);
        updateLastInsertedIndices();
        return this;
    }

    @Override
    public IManipulacaoDados inserirDados(List<String> dados, String delimitador) {
        insersorDeDados.inserirDados(dados, delimitador);
        updateLastInsertedIndices();
        return this;
    }

    @Override
    public IManipulacaoDados inserirDadosArquivo(String caminhoArquivo, String delimitador) {
        insersorDeDados.inserirDadosArquivo(caminhoArquivo, delimitador);
        updateLastInsertedIndices();
        return this;
    }

    @Override
    public IManipulacaoDados converterEmNumero(String posicaoInicial) {
        Conversores.converterEmNumero(sheet, posicaoInicial);
        return this;
    }

    @Override
    public IManipulacaoDados converterEmContabil(String posicaoInicial) {
        Conversores.converterEmContabil(sheet, posicaoInicial, workbook);
        return this;
    }

    @Override
    public IManipulacaoDados somarColuna(String posicaoInicial) {
        Calculos.somarColuna(sheet, posicaoInicial);
        String colunaLetra = posicaoInicial.replaceAll("[0-9]", "");
        this.ultimaLinha(colunaLetra);
        ultimoIndiceDeColunaInserido = -1;
        return this;
    }

    @Override
    public IManipulacaoDados somarColunaComTexto(String posicaoInicial, String texto) {
        Calculos.somarColunaComTexto(sheet, posicaoInicial, texto);
        String colunaLetra = posicaoInicial.replaceAll("[0-9]", "");
        this.ultimaLinha(colunaLetra);
        ultimoIndiceDeColunaInserido = -1;
        return this;
    }

    /**
     * Atualiza os índices da última célula inserida.
     */
    private void updateLastInsertedIndices() {
        ultimoIndiceDeLinhaInserido = insersorDeDados.getUltimoIndiceDeLinhaInserido();
        ultimoIndiceDeColunaInserido = insersorDeDados.getUltimoIndiceDeColunaInserido();
    }

    // Métodos de IEstilos

    @Override
    public EstiloCelula aplicarEstilos() {
        EstiloCelula estilo;

        if (positionManager.isTodaPlanilhaDefinida()) {
            estilo = new EstiloCelula(workbook, sheet);
        } else if (positionManager.isIntervaloDefinida()) {
            estilo = new EstiloCelula(workbook, sheet, positionManager.getPosicaoInicialLinha(),
                    positionManager.getPosicaoInicialColuna(), positionManager.getPosicaoFinalLinha(),
                    positionManager.getPosicaoFinalColuna());
        } else if (ultimoIndiceDeLinhaInserido == -1) {
            estilo = new EstiloCelula(workbook, sheet, -1, -1);
        } else {
            estilo = new EstiloCelula(workbook, sheet, ultimoIndiceDeLinhaInserido, -1);
        }

        positionManager.resetarPosicao();
        return estilo;
    }

    @Override
    public EstiloCelula centralizarTudo() {
        return aplicarEstilos().centralizarTudo();
    }

    @Override
    public EstiloCelula redimensionarColunas() {
        return aplicarEstilos().redimensionarColunas();
    }

    @Override
    public EstiloCelula removerLinhasDeGrade() {
        return aplicarEstilos().removerLinhasDeGrade();
    }

    @Override
    public EstiloCelula aplicarEstilosEmCelula() {
        if (ultimoIndiceDeLinhaInserido == -1 || ultimoIndiceDeColunaInserido == -1) {
            return new EstiloCelula(workbook, sheet, -1, -1);
        }
        return new EstiloCelula(workbook, sheet, ultimoIndiceDeLinhaInserido, ultimoIndiceDeColunaInserido);
    }

    @Override
    public EstiloCelula todasAsBordasEmTudo() {
        // Implementação real para aplicar bordas
        aplicarEstilos().aplicarBordasEspessasComInternas("A1", "Z100");
        return aplicarEstilos();
    }
}
