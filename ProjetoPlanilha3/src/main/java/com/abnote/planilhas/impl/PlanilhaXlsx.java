package com.abnote.planilhas.impl;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.abnote.planilhas.estilos.EstiloCelula;
import com.abnote.planilhas.interfaces.IManipulacaoDados;

public class PlanilhaXlsx extends PlanilhaBase {

    @Override
    protected void inicializarWorkbook() {
        workbook = new XSSFWorkbook(); // Inicializa XSSFWorkbook para arquivos .xlsx
    }

    @Override
    public EstiloCelula todasAsBordasEmTudo() {
        // Implementação específica para bordas em todos os cantos
        return super.todasAsBordasEmTudo();
    }

	@Override
	public IManipulacaoDados inserir(String valor) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public IManipulacaoDados inserir(int valor) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public IManipulacaoDados inserir(double valor) {
		// TODO Auto-generated method stub
		return null;
	}

    // Caso haja métodos específicos para .xlsx, eles podem ser implementados aqui
}
