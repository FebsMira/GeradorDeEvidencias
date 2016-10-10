package br.com.felipemira.arquivos.office;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import br.com.felipemira.arquivos.office.custom.document.CustomXWPFDocument;
import br.com.felipemira.arquivos.office.iterator.WordIterator;
import br.com.felipemira.convert.ConvertToPDF;
import br.com.felipemira.copy.Copy;

public class WordAutomacao {
	
	private String nomeArquivo;
	private String nomeComHora;
	private String horaExecucao;
	private Calendar horaInicio;
	private String caminhoTemplate;
	private String caminhoImagem;
	private String caminhoDoc;
	private int countFalhou = 0;
	
	/**
	 * Cria um documento Word através do template.
	 * @param nomeArquivo - String.
	 * @param caminhoDoc - String.
	 * @param caminhoTemplate - String.
	 * @param caminhoImagem - String.
	 */
	public WordAutomacao(String nomeArquivo, String caminhoDoc, String caminhoTemplate, String caminhoImagem){
		
		this.nomeArquivo = nomeArquivo;
		this.caminhoDoc = caminhoDoc;
		this.caminhoTemplate = caminhoTemplate;
		this.caminhoImagem = caminhoImagem;
		
		try {
			criarDocumento();
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
	}
	
	/**
	 * Pega o templete da evidência e cria um novo documento word.
	 * @throws IOException
	 */
	private void criarDocumento() throws IOException{
		//Pega a hora que foi criado o documento.
		DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
        Calendar cal = Calendar.getInstance();
        this.horaInicio = cal;
        this.horaExecucao = dateFormat.format(cal.getTime());
        String date = this.horaExecucao.replaceAll("/", "_").replaceAll(" ", "_").replaceAll(":", "_");
        
        //Implementa o nome do arquivo de evidencias que sera criado.
        this.nomeComHora = nomeArquivo + "_" + date;
        
        //Defino para toda a classe o caminho da minha evidencia
        this.caminhoDoc = this.caminhoDoc + this.nomeComHora + ".docx";
        
        
		CustomXWPFDocument document = null;
		
        //Crio um file com o template.
        File file = new File(this.caminhoTemplate);
        
        //Faco uma copia do template renomeada com o nome da minha evidencia
        if(file.exists()) {
        	File copiaFile = new File(caminhoDoc);
        	Copy.copyFile(file, copiaFile);
        	InputStream arquivo = new FileInputStream(caminhoDoc);
	        @SuppressWarnings("resource")
			CustomXWPFDocument copiaDocumento = new CustomXWPFDocument(arquivo);
	        document = copiaDocumento;
        }else{
        	System.out.println("templateEvidencia.docx não consta na pasta informada!");
        }
        
        //Insere o nome do CT na tabela um, primeira linha da primeira coluna.
        WordIterator.inserirDadoTabela(this.caminhoDoc, document, 0, 0, 0, this.nomeArquivo);
        //Insere a hora de execução na tabela um, primeira linha da segunda coluna.
        WordIterator.inserirDadoTabela(this.caminhoDoc, document, 0, 0, 1, this.horaExecucao);
        
        //Insere o título no documento word.
        WordIterator.inserirTitulo(this.caminhoDoc, document, this.nomeArquivo);
        
	}
	
	
	/**
	 * Insere uma evidência no arquivo Word. Para gerar imagem é necessário passar true no parâmetro imagem.
	 * @param mensagem - String
	 * @param passouFalhou - Boolean
	 * @param imagem - Boolean
	 */
	public void inserirEvidencia(String mensagem, Boolean passouFalhou, Boolean imagem){
		if(passouFalhou){
			mensagem = mensagem + " - Passou";
		}else{
			mensagem = mensagem + " - Falhou";
			this.countFalhou = 1;
		}
		try {
			gerarEvidencia(mensagem, imagem);
		} catch (InvalidFormatException e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		} catch (AWTException e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}
	
	/**
	 * Gera uma evidencia e insere no arquivo Word.
	 * @param mensagem - String
	 * @param geraImagem - Boolean
	 * @throws IOException
	 * @throws AWTException
	 * @throws InvalidFormatException
	 */
	private void gerarEvidencia(String mensagem, Boolean geraImagem) throws IOException, AWTException, InvalidFormatException{
		CustomXWPFDocument document = null;
        InputStream arquivo = new FileInputStream(this.caminhoDoc);
		CustomXWPFDocument copiaDocumento = new CustomXWPFDocument(arquivo);
	    document = copiaDocumento;
       
	    if(geraImagem){
	    	WordIterator.inserirTexto(this.caminhoDoc, document, mensagem);
	    	WordIterator.inserirImagem(this.caminhoDoc, document, this.caminhoImagem);
	    }else{
	    	WordIterator.inserirTexto(this.caminhoDoc, document, mensagem);
	    }
	}
	
	/**
	 * Calcula a duração da execução.
	 * @return String.
	 */
	private String calcularDuracao(){
        
        long diferenca = System.currentTimeMillis() - this.horaInicio.getTimeInMillis();
        long diferencaHoras = diferenca / (60 * 60 * 1000);
        long diferencaMin = (diferenca % (60 * 60 * 1000))/ (60 * 1000);
        long diferencaSeg = (diferenca % (60 * 60 * 1000))% (60 * 1000) / 1000;
        
        String horas = "";
        String minutos = "";
        String segundos = "";
        
        if(String.valueOf(diferencaHoras).length() == 1){
        	horas = "0" + String.valueOf(diferencaHoras);
        }else{
        	horas = String.valueOf(diferencaHoras);
        }
        
        if(String.valueOf(diferencaMin).length() == 1){
        	minutos = "0" + String.valueOf(diferencaMin);
        }else{
        	minutos = String.valueOf(diferencaMin);
        }
        
        if(String.valueOf(diferencaSeg).length() == 1){
        	segundos = "0" + String.valueOf(diferencaSeg);
        }else{
        	segundos = String.valueOf(diferencaSeg);
        }
        
        
        return horas + ":" + minutos + ":" + segundos;
	}
	
	/**
	 * Finaliza a evidência e cria o arquivo PDF
	 * @throws IOException
	 */
	public void finalizarEvidencia() throws IOException{
		CustomXWPFDocument document = null;
        InputStream arquivo = new FileInputStream(this.caminhoDoc);
		CustomXWPFDocument copiaDocumento = new CustomXWPFDocument(arquivo);
	    document = copiaDocumento;
	    
	    //Insere a duração do CT na tabela um, segunda linha, segunda coluna.
	    WordIterator.inserirDadoTabela(this.caminhoDoc, document, 0, 1, 1, calcularDuracao());
	    
	    //Insere o Status do caso de teste na tabela.
	    if(this.countFalhou == 1){
	    	WordIterator.inserirDadoTabela(this.caminhoDoc, document, 0, 1, 0, "Falhou");
	    }else{
	    	WordIterator.inserirDadoTabela(this.caminhoDoc, document, 0, 1, 0, "Passou");
	    }
	    
	    ConvertToPDF.convert(this.caminhoDoc, this.nomeComHora);
	}
}
