package br.com.felipemira;

import java.io.IOException;

import br.com.felipemira.arquivos.office.WordAutomacao;

public class RftToWord {

	public static void main(String[] args) throws InterruptedException, IOException {
		WordAutomacao word = new WordAutomacao("Evidencia"
				, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Evidencias\\"
				, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\Template.docx"
				, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\evidencia.png");
		
		Thread.sleep(5000);
		
		word.inserirEvidencia("Teste", true, true);
		
		word.finalizarEvidencia();
	}

}
