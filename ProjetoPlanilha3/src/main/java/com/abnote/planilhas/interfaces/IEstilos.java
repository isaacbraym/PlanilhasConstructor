package com.abnote.planilhas.interfaces;

import com.abnote.planilhas.estilos.EstiloCelula;

public interface IEstilos {

    EstiloCelula aplicarEstilos();

    EstiloCelula centralizarTudo();

    EstiloCelula redimensionarColunas();

    EstiloCelula removerLinhasDeGrade();

    EstiloCelula aplicarEstilosEmCelula();

    EstiloCelula todasAsBordasEmTudo();
}
