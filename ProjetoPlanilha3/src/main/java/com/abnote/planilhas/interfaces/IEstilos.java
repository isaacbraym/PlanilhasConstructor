package com.abnote.planilhas.interfaces;

import org.apache.poi.ss.usermodel.Cell;

import com.abnote.planilhas.estilos.EstiloCelula;

public interface IEstilos {

	EstiloCelula centralizarTudo();

	EstiloCelula redimensionarColunas();

	EstiloCelula removerLinhasDeGrade();

	EstiloCelula aplicarEstilosEmCelula();

	EstiloCelula todasAsBordasEmTudo();

}
