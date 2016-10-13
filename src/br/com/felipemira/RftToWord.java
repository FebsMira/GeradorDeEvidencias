package br.com.felipemira;

import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.swing.JOptionPane;

import br.com.felipemira.arquivos.office.WordAutomacao;
import br.com.felipemira.arquivos.office.reader.ExcelReader;
import br.com.felipemira.objects.object.CasoDeTeste;

public class RftToWord {

	public static void executar(String caminhoSalvar, String caminhoModelo, String caminhoExcel, int linhaInicio, int linhaFim) throws InterruptedException, IOException {
		
		Map<Integer, CasoDeTeste> listaCasosDeTeste = new HashMap<Integer, CasoDeTeste>();
		
		//ExcelReader leitorRTF = new ExcelReader("C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\RTFs\\RTF - Nova Consulta - Exemplos.xlsx");
		ExcelReader leitorRTF = new ExcelReader(caminhoExcel);
		
		//listaCasosDeTeste = leitorRTF.readerDelimited(12, 21);
		//listaCasosDeTeste = leitorRTF.readerDelimited(12, 57);
		//listaCasosDeTeste = leitorRTF.readerDelimited(12, 1325);
		listaCasosDeTeste = leitorRTF.readerDelimited(linhaInicio, linhaFim);
		
		@SuppressWarnings("rawtypes")
		Iterator iterator = listaCasosDeTeste.keySet().iterator();
		
		while(iterator.hasNext()){
			CasoDeTeste casoDeTeste = (CasoDeTeste) listaCasosDeTeste.get(iterator.next());
			
			/*WordAutomacao word = new WordAutomacao(casoDeTeste.getSiglaCasoDeTeste() + " - " + casoDeTeste.getIdCasoDeTeste()
					, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Evidencias\\"
					, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\Modelo "+ casoDeTeste.getProcedimentosDeExecucao().size() +".docx"
					, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\evidencia.png", casoDeTeste);*/
			
			WordAutomacao word = new WordAutomacao(casoDeTeste.getSiglaCasoDeTeste() + " - " + casoDeTeste.getIdCasoDeTeste()
			, caminhoSalvar + "\\"
			, caminhoModelo + "\\Modelo " + casoDeTeste.getProcedimentosDeExecucao().size() +".docx"
			, caminhoModelo + "\\evidencia.png", casoDeTeste);
			
			System.out.println("\nCenário Caso de teste: " + casoDeTeste.getCenarioDeTeste() + "\n");
			
			Map<Integer, String> procedimentosDeExecucao = casoDeTeste.getProcedimentosDeExecucao();
			Map<Integer, String> resultadosEsperados = casoDeTeste.getResultadoEsperado();
			
			@SuppressWarnings("rawtypes")
			Iterator iteratorProcedimentos = procedimentosDeExecucao.keySet().iterator();
			
			System.out.println("Procedimentos: \n");
			while(iteratorProcedimentos.hasNext()){
				int numeroIterator = (int) iteratorProcedimentos.next();
				word.inserirDadoTabela(numeroIterator, 0, 1, "Passo " + numeroIterator);
				word.inserirDadoTabela(numeroIterator, 1, 1, (String) procedimentosDeExecucao.get(numeroIterator));
				word.inserirDadoTabela(numeroIterator, 2, 1, (String) resultadosEsperados.get(numeroIterator));
				String procedimento = (String) procedimentosDeExecucao.get(numeroIterator);
				System.out.println(procedimento);
			}
		}
		
		JOptionPane.showMessageDialog(null, "Criação de documentos referentes a RTF finalizado!", "Finalizado!", JOptionPane.INFORMATION_MESSAGE);
	}

}
