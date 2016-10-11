package br.com.felipemira;

import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import br.com.felipemira.arquivos.office.reader.ExcelReader;
import br.com.felipemira.objects.object.CasoDeTeste;

public class RftToWord {

	public static void main(String[] args) throws InterruptedException, IOException {
		
		Map<Integer, CasoDeTeste> listaCasosDeTeste = new HashMap<Integer, CasoDeTeste>();
		
		ExcelReader leitorRTF = new ExcelReader("C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\RTFs\\RTF - Nova Consulta - Exemplos.xlsx");
		
		listaCasosDeTeste = leitorRTF.readerDelimited(12, 22);
		
		@SuppressWarnings("rawtypes")
		Iterator iterator = listaCasosDeTeste.keySet().iterator();
		
		while(iterator.hasNext()){
			CasoDeTeste casoDeTeste = (CasoDeTeste) listaCasosDeTeste.get(iterator.next());
			System.out.println(casoDeTeste.getCenarioDeTeste());
		}
		
		//leitorRTF.readerDelimited(12, 12, 13, 13);
		
		/*WordAutomacao word = new WordAutomacao("Evidencia"
				, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Evidencias\\"
				, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\Template.docx"
				, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\evidencia.png");
		
		Thread.sleep(5000);
		
		word.inserirEvidencia("Teste", true, true);
		
		word.finalizarEvidencia();*/
	}

}
