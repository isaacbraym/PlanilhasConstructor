package com.abnote.planilhas.impl;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.abnote.planilhas.interfaces.IPlanilha;

public class PlanilhaXlsx extends PlanilhaBase {
	@Override
	protected void inicializarWorkbook() {
		workbook = new XSSFWorkbook(); // Inicializa XSSFWorkbook para arquivos .xlsx
	}

	@Override
	public IPlanilha todasAsBordasEmTudo() {
		// TODO Auto-generated method stub
		return null;
	}
}