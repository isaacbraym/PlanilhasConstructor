package br.com.abnote.scp2.base.valid.leitordeplanilhas.util;

import java.util.logging.*;

import com.abnote.planilhas.utils.ColorFormatter;

public class LoggerUtil {

    static {
        // Configurar o logger global
        Logger rootLogger = Logger.getLogger("");

        // Remover os handlers padrão
        Handler[] handlers = rootLogger.getHandlers();
        if (handlers != null) {
            for (Handler handler : handlers) {
                rootLogger.removeHandler(handler);
            }
        }

        // Criar um ConsoleHandler personalizado que escreve em System.out
        ConsoleHandler consoleHandler = new ConsoleHandler() {
            {
                setOutputStream(System.out);
            }
        };

        // Definir o nível do handler e do logger
        consoleHandler.setLevel(Level.ALL); // Capturar todos os níveis
        rootLogger.setLevel(Level.ALL);

        // Definir o ColorFormatter para o handler
        consoleHandler.setFormatter(new ColorFormatter());

        // Adicionar o handler ao logger raiz
        rootLogger.addHandler(consoleHandler);
    }

    // Método utilitário para obter o logger configurado
    public static Logger getLogger(Class<?> clazz) {
        return Logger.getLogger(clazz.getName());
    }
}
