package com.abnote.planilhas.impl;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.interfaces.IEstilos;
import com.abnote.planilhas.utils.PositionManager;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class StyleManager implements IEstilos {

	private final Workbook workbook;
	private final Sheet sheet;
	private final PositionManager positionManager;
	private final DataManipulator dataManipulator;

	public StyleManager(Workbook workbook, Sheet sheet, PositionManager positionManager,
			DataManipulator dataManipulator) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.positionManager = positionManager;
		this.dataManipulator = dataManipulator;
	}

	@Override
	public EstiloCelula aplicarEstilos() {
		EstiloCelula estilo;

		if (positionManager.isTodaPlanilhaDefinida()) {
			estilo = new EstiloCelula(workbook, sheet);
		} else if (positionManager.isIntervaloDefinida()) {
			estilo = new EstiloCelula(workbook, sheet, positionManager.getPosicaoInicialLinha(),
					positionManager.getPosicaoInicialColuna(), positionManager.getPosicaoFinalLinha(),
					positionManager.getPosicaoFinalColuna());
		} else if (dataManipulator.getUltimoIndiceDeLinhaInserido() == -1) {
			estilo = new EstiloCelula(workbook, sheet, -1, -1);
		} else {
			estilo = new EstiloCelula(workbook, sheet, dataManipulator.getUltimoIndiceDeLinhaInserido(), -1);
		}

		// Resetar o positionManager aqui
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
		if (dataManipulator.getUltimoIndiceDeLinhaInserido() == -1
				|| dataManipulator.getUltimoIndiceDeColunaInserido() == -1) {
			return new EstiloCelula(workbook, sheet, -1, -1);
		}
		return new EstiloCelula(workbook, sheet, dataManipulator.getUltimoIndiceDeLinhaInserido(),
				dataManipulator.getUltimoIndiceDeColunaInserido());
	}

	@Override
	public EstiloCelula todasAsBordasEmTudo() {
		// Implementação real para aplicar bordas
		aplicarEstilos().aplicarBordasEspessasComInternas("A1", "Z100");
		return aplicarEstilos();
	}
}
