package com.abnote.planilhas.estilos.estilos;

public enum CorEnum {
    VERMELHO_ESCURO(139, 0, 0),
    AZUL(0, 0, 255),
    VERDE(0, 128, 0),
    PRETO(0, 0, 0),
    BRANCO(255, 255, 255),
    CINZA_CLARO(211, 211, 211),
    AMARELO(255, 255, 0),
    LARANJA(255, 165, 0),
    ROSA(255, 192, 203),
    ROXO(128, 0, 128),
    VIOLETA(238, 130, 238),
    AZUL_CELESTE(135, 206, 235),
    VERDE_LIMAO(50, 205, 50),
    MARROM(165, 42, 42),
    DOURADO(255, 215, 0),
    PRATA(192, 192, 192),
    BEGE(245, 245, 220),
    SALMAO(250, 128, 114),
    TURQUESA(64, 224, 208),
    LAVANDA(230, 230, 250),
    AZUL_MARINHO(0, 0, 128),
    CINZA_ESCURO(169, 169, 169),
    OLIVA(128, 128, 0),
    CORAL(255, 127, 80),
    MENTA(189, 252, 201),
    VERDE_OLIVA(107, 142, 35),
    VERDE_AGUA(32, 178, 170);

    private final int red;
    private final int green;
    private final int blue;

    CorEnum(int red, int green, int blue) {
        this.red = red;
        this.green = green;
        this.blue = blue;
    }

    public int getRed() {
        return red;
    }

    public int getGreen() {
        return green;
    }

    public int getBlue() {
        return blue;
    }
}
