package br.com.felipemira;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import br.com.felipemira.arquivos.office.WordAutomacao;
import br.com.felipemira.arquivos.office.reader.ExcelReader;
import br.com.felipemira.objects.object.CasoDeTeste;
import br.com.felipemira.objects.object.Error;

public class RftToWord {

	public static void executar(String caminhoSalvar, String caminhoModelo, String caminhoExcel, int linhaInicio, int linhaFim) {
		
		Map<Integer, CasoDeTeste> listaCasosDeTeste = new HashMap<Integer, CasoDeTeste>();
		
		//ExcelReader leitorRTF = new ExcelReader("C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\RTFs\\RTF - Nova Consulta - Exemplos.xlsx");
		ExcelReader leitorRTF = new ExcelReader(caminhoExcel);
		
		//listaCasosDeTeste = leitorRTF.readerDelimited(12, 21);
		//listaCasosDeTeste = leitorRTF.readerDelimited(12, 57);
		//listaCasosDeTeste = leitorRTF.readerDelimited(12, 1325);
		try {
			listaCasosDeTeste = leitorRTF.readerDelimited(linhaInicio, linhaFim);
		} catch (Exception e) {
			Error.error = e.getMessage().toString();
		}
		
		@SuppressWarnings("rawtypes")
		Iterator iterator = listaCasosDeTeste.keySet().iterator();
		
		while(iterator.hasNext()){
			CasoDeTeste casoDeTeste = (CasoDeTeste) listaCasosDeTeste.get(iterator.next());
			
			/*WordAutomacao word = new WordAutomacao(casoDeTeste.getSiglaCasoDeTeste() + " - " + casoDeTeste.getIdCasoDeTeste()
					, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Evidencias\\"
					, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\Modelo "+ casoDeTeste.getProcedimentosDeExecucao().size() +".docx"
					, "C:\\Users\\Felipe Mira\\Documents\\WordToPDF\\Template\\evidencia.png", casoDeTeste);*/
			
			WordAutomacao word = new WordAutomacao(casoDeTeste.getSiglaCasoDeTeste() + " - " + casoDeTeste.getIdCasoDeTeste()
			, caminhoSalvar + "/"
			, caminhoModelo + "/Modelo " + /*casoDeTeste.getProcedimentosDeExecucao().size()*/ "Mira.docx"
			, caminhoModelo + "/evidencia.png", casoDeTeste);
			
			System.out.println("\nCenario Caso de teste: " + casoDeTeste.getCenarioDeTeste() + "\n");
			
			Map<Integer, String> preCondicoes = casoDeTeste.getPreCondicoes();
			Map<Integer, String> procedimentosDeExecucao = casoDeTeste.getProcedimentosDeExecucao();
			Map<Integer, String> resultadosEsperados = casoDeTeste.getResultadoEsperado();
			
			@SuppressWarnings("rawtypes")
			Iterator iteratorProcedimentos = procedimentosDeExecucao.keySet().iterator();
			
			System.out.println("Procedimentos: \n");
			while(iteratorProcedimentos.hasNext()){
				int numeroIterator = (int) iteratorProcedimentos.next();
				
				if(numeroIterator > 1){
					word.newPage();
				}
				word.newTableFuncional(numeroIterator);
				
				//Insere numero do passo.
				word.inserirDadoTabela(numeroIterator, 0, 0, numeroIterator + ":", true);
				
				//Procedimentos para execucao do passo.
				word.inserirDadoTabela(numeroIterator, 0, 1, (String) procedimentosDeExecucao.get(numeroIterator), false);
				
				//Pre Condicoes
				word.inserirDadoTabela(numeroIterator, 1, 1, "Pré-Condiçoes: ", true);
				word.inserirDadoTabela(numeroIterator, 1, 1, (preCondicoes.get(numeroIterator) == null)? " " : (String) preCondicoes.get(numeroIterator), false);
				
				//Resultado esperado do passo.
				word.inserirDadoTabela(numeroIterator, 2, 1, (String) resultadosEsperados.get(numeroIterator), false);
				
				//Apenas para fins de teste.
				String procedimento = (String) procedimentosDeExecucao.get(numeroIterator);
				System.out.println(procedimento);
			}
		}
	}
}
