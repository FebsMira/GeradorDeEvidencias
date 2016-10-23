package br.com.felipemira.arquivos.office.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import br.com.felipemira.objects.iterator.CasoDeTesteIterator;
import br.com.felipemira.objects.object.CasoDeTeste;

public class ExcelReader {
	
	private static String caminhoExcel;
	
	/**
	 * Recebe uma planilha excel, que ser� preparada para a leitura.
	 * @param caminhoExcel - String.
	 */
	public ExcelReader(String caminhoExcel){
		ExcelReader.caminhoExcel = caminhoExcel;
	}
	
	/**
	 * L� o excel apartir dos par�metros passados. Deve receber a linha de inicio e a linha final para leitura, assim como a coluna de in�cio e a coluna final.
	 * @param linhaInicio - int.
	 * @param linhaFim - int.
	 * @throws IOException
	 */
	public Map<Integer, CasoDeTeste> readerDelimited(int linhaInicio, int linhaFim) throws IOException{
		//inst�ncia um Map que retornar� para a classe que chamou o m�todo os casos de teste que o Excel cont�m, nas linhas delimitadas.
		Map<Integer, CasoDeTeste> listaCasosDeTeste = new HashMap<Integer, CasoDeTeste>();
		
		//Leitor de arquivos recebe o excel informado.
		FileInputStream inputStream = new FileInputStream(new File(caminhoExcel));
		
		//O Workbook do POI recebe o excel.
		Workbook workbook = new XSSFWorkbook(inputStream);
		//Aba do excel a ser selecionada.
		Sheet sheet = workbook.getSheetAt(0);
		
		//A lista de casos de teste recebe o Map com todos os casos de teste lidos.
		listaCasosDeTeste = readerRows(workbook, sheet, linhaInicio, linhaFim);
		
		workbook.close();
		inputStream.close();
		
		return listaCasosDeTeste;
	}
	
	/**
	 * Percorre o excel delimitado por linha e coluna, e chama o m�todo readerCell passando as colunas que devem ser lidas.
	 * @param workbook - Workbook.
	 * @param sheet - Sheet.
	 * @param linhaInicio - int.
	 * @param linhaFim - int.
	 */
	private static Map<Integer, CasoDeTeste> readerRows(Workbook workbook, Sheet sheet, int linhaInicio, int linhaFim){
		Sheet firstSheet = sheet;
		
		Row row;
		
		//Inicio das colunas do Caso de Teste
		int colunaInicio = 2;
		//Fim das colunas do Caso de Teste
		int colunaFim = 13;
		
		@SuppressWarnings("rawtypes")
		Iterator iterator = firstSheet.rowIterator();
		
		int countLinha = 0;
		
		//Inst�ncia de um objeto CasoDeTeste, para ser preenchida e adicionada a lista.
		CasoDeTeste casoDeTeste = new CasoDeTeste();
		//Recebe o neg�cio da c�lula D7 do excel.
		String negocio = readerRow(workbook, firstSheet, 7, 4);
		//Insere o neg�cio no caso de teste.
		casoDeTeste.setNegocio(negocio);
		
		
		//Contador respons�vel por contar a quantidade dos casos de testes e de atribuir a chave para o Map.
		int countCasosDeTeste = 0;
		
		//Lista que receber� os casos de testes.
		Map<Integer, CasoDeTeste> listaCasosDeTeste = new HashMap<Integer, CasoDeTeste>();
		
		//Percorre as linhas informadas.
		while((iterator.hasNext() && countLinha == 0) || (iterator.hasNext() && countLinha <= (linhaFim - 1))){
			row = (Row) iterator.next();
			
			if(countLinha >= (linhaInicio-1) && countLinha <= (linhaFim - 1)){
				//readerCells(row, colunaInicio, colunaFim);
				
				//L� cada uma das c�lulas da linha, se n�o encontrar nada na c�lula 8 ou 9 � sinalizado que acabou os passos do caso de teste.
				if(readerCellsDelimited(casoDeTeste, row, colunaInicio, colunaFim) || row.getRowNum() == (linhaFim - 1)){
					//Se o Caso de Teste chegou ao fim implementa o contador, que � o controle para adi��o de elementos no Map.
					countCasosDeTeste ++;
					//Preenche o Map com o caso de teste, com a Key do contador.
					listaCasosDeTeste.put(countCasosDeTeste, casoDeTeste);
					//Cria uma nova inst�ncia para o objeto casoDeTeste.
					casoDeTeste = new CasoDeTeste();
					//Insere o neg�cio no caso de teste.
					casoDeTeste.setNegocio(negocio);
				}
			}
			countLinha ++;
		}
		return listaCasosDeTeste;
	}
	
	/**
	 * L� e grava os par�metros dentro do objeto caso de teste. Retorna true ao chegar no final do caso de teste.
	 * @param casoDeTeste - CasoDeTeste.
	 * @param row - Row.
	 * @param colunaInicio int.
	 * @param colunaFim - int.
	 * @return boolean - se acabou o caso de teste retorna true;
	 */
	private static boolean readerCellsDelimited(CasoDeTeste casoDeTeste, Row row, int colunaInicio, int colunaFim){
		//Vari�vel que ir� controlar se o caso de teste chegou ao fim.
		boolean finalCasoDeTeste = false;
		
		//Adicionado para pegar a coluna D7 onde cont�m o neg�cio relacionado a RTF
		
		
		for(int i = colunaInicio - 1; i < colunaFim; i++){
			if(row.getCell(i) != null){
				
				if(row.getCell(i).getCellType() == Cell.CELL_TYPE_STRING){
					
					//Se o elemento que est� sendo puxado do excel n�o for vazio ele adicionar� o elemento ao caso de teste passado na chamada do m�todo.
					if(!row.getCell(i).getStringCellValue().equals("")){
						CasoDeTesteIterator.gravarDados(casoDeTeste, i, row.getCell(i).getStringCellValue());
					}
					
				}else if(row.getCell(i).getCellType() == Cell.CELL_TYPE_NUMERIC){
					
					//Se o elemento que est� sendo puxado do excel n�o for vazio ele adicionar� o elemento ao caso de teste passado na chamada do m�todo.
					if(!String.valueOf(row.getCell(i).getNumericCellValue()).equals("")){
						CasoDeTesteIterator.gravarDados(casoDeTeste, i, String.valueOf(row.getCell(i).getNumericCellValue()));
					}
				
				//Se o elemento for em branco ir� verificar os �ndices 8 e 9.
				}else if(row.getCell(i).getCellType() == Cell.CELL_TYPE_BLANK){
					switch(i){
						case 1:
						case 2:
						case 3:
						case 4:
						case 5:
						case 6:
						case 7:
							break;
						//Se for a c�lula com �ndice 8 ou 9, que sinalizam os passos e resultados do caso de teste, sinaliza para a fun��o que chamou que o Caso de Teste chegou ao fim.
						case 8:
						case 9:
							finalCasoDeTeste = true;
							break;
						case 10:
						case 11:
						case 12:
							break;
					}	
				}
				
			//Para efeito de c�lulas null, o leitor passar� sobre elas sem nenhum problema.
			}else{
				switch(i){
					case 1:
					case 2:
					case 3:
					case 4:
					case 5:
					case 6:
					case 7:
						break;
					//Se for a c�lula com �ndice 8 ou 9, que sinalizam os passos e resultados do caso de teste, sinaliza para a fun��o que chamou que o Caso de Teste chegou ao fim.
					case 8:
					case 9:
						finalCasoDeTeste = true;
						break;
					case 10:
					case 11:
					case 12:
						break;
				}	
			}
		}
		
		//Retorna se o caso de teste chegou ao fim.
		if(finalCasoDeTeste){
			return true;
		}else{
			return false;
		}
	}
	
	/**
	 * Percorre o excel para pegar o conte�do de uma �nica c�lula em uma �nica coluna.
	 * @param workbook - Workbook
	 * @param linha - Int
	 * @param coluna -Int
	 * @return String - Contendo dado.
	 */
	private static String readerRow(Workbook workbook, Sheet sheet, int linha, int coluna){
		String negocio = null;
		Sheet firstSheet = sheet;
		
		Row row;
		
		@SuppressWarnings("rawtypes")
		Iterator iterator = firstSheet.rowIterator();
		
		int countLinha = 0;
		
		//Percorre as linhas informadas.
		while((iterator.hasNext() && countLinha <= (linha-1) )){
			row = (Row) iterator.next();
			
			if(countLinha == (linha-1)){
				negocio = readerCell(row, coluna);
			}
			countLinha ++;
		}
		return negocio;
	}
	
	/**
	 * Leitor de uma c�lula do Excel.
	 * @param row - Row.
	 * @param colunaInicio - int.
	 * @return String - contendo dado da c�lula.
	 */
	private static String readerCell(Row row, int coluna){
		String dado = null;
		
		coluna = coluna - 1;
			
		if(row.getCell(coluna).getCellType() == Cell.CELL_TYPE_STRING){
			
			//Se o elemento que est� sendo puxado do excel n�o for vazio ele guardar� na String.
			if(!row.getCell(coluna).getStringCellValue().equals("")){
				dado =  row.getCell(coluna).getStringCellValue();
			}
			
		}else if(row.getCell(coluna).getCellType() == Cell.CELL_TYPE_NUMERIC){
			
			//Se o elemento que est� sendo puxado do excel n�o for vazio ele guardar� na String.
			if(!String.valueOf(row.getCell(coluna).getNumericCellValue()).equals("")){
				dado = String.valueOf(row.getCell(coluna).getNumericCellValue());
			}
		}
		return dado;
	}
}
