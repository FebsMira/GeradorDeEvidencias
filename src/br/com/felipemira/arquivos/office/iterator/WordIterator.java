package br.com.felipemira.arquivos.office.iterator;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import br.com.felipemira.arquivos.office.custom.document.CustomXWPFDocument;

public class WordIterator {
	
	/**
	 * Insere dados em um cabeçalho 2x2 no template word.
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param titulo - String
	 * @param horaExecucao - String
	 * @throws IOException
	 */
	public static void inserirCabecalho(String caminhoDoc, CustomXWPFDocument documento, String titulo, String horaExecucao) throws IOException{
		XWPFTable table = documento.getTableArray(0);
		
		XWPFRun casoDeTeste = table.getRow(0).getCell(0).getParagraphs().get(0).createRun();
		casoDeTeste.setBold(false);
		casoDeTeste.setFontSize(10);
		casoDeTeste.setFontFamily("Verdana");
		casoDeTeste.setColor("000000");
		casoDeTeste.setText(" " + titulo);
		
		XWPFRun horaDeExecucao = table.getRow(0).getCell(1).getParagraphs().get(0).createRun();
		horaDeExecucao.setBold(false);
		horaDeExecucao.setFontSize(10);
		horaDeExecucao.setFontFamily("Verdana");
		horaDeExecucao.setColor("000000");
		horaDeExecucao.setText(" " + horaExecucao);
		
		FileOutputStream outStream = null;
        outStream = new FileOutputStream(caminhoDoc);
        
        documento.write(outStream);
        outStream.close();
		
	}
	
	/**
	 * Insere a duracao da execução no arquivo word.
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param duracao - String
	 * @throws IOException
	 */
	public static void inserirDuracao(String caminhoDoc, CustomXWPFDocument documento, String duracao) throws IOException{
		XWPFTable table = documento.getTableArray(0);
		
		XWPFRun casoDeTeste = table.getRow(1).getCell(1).getParagraphs().get(0).createRun();
		casoDeTeste.setBold(false);
		casoDeTeste.setFontSize(10);
		casoDeTeste.setFontFamily("Verdana");
		casoDeTeste.setColor("000000");
		casoDeTeste.setText(" " + duracao);
		
		FileOutputStream outStream = null;
        outStream = new FileOutputStream(caminhoDoc);
        
        documento.write(outStream);
        outStream.close();
	}
	
	/**
	 * Insere um Titulo no arquivo word
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param titulo - String
	 * @throws IOException
	 */
	public static void inserirTitulo(String caminhoDoc, CustomXWPFDocument documento, String titulo) throws IOException{
		XWPFParagraph paragraphOne = documento.createParagraph();
        paragraphOne.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun paragraphOneRunOne = paragraphOne.createRun();
        paragraphOneRunOne.setBold(true);
        paragraphOneRunOne.setFontSize(14);
        paragraphOneRunOne.setFontFamily("Verdana");
        paragraphOneRunOne.setColor("000000");
        
        paragraphOneRunOne.setText(titulo);
        
        FileOutputStream outStream = null;
        outStream = new FileOutputStream(caminhoDoc);
        
        documento.write(outStream);
        outStream.close();
	}
	
	/**
	 * Insere um texto no documento word.
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param texto - String
	 * @throws IOException
	 */
	public static void inserirTexto(String caminhoDoc, CustomXWPFDocument documento, String texto) throws IOException{
		XWPFParagraph paragraphFour = documento.createParagraph();
        paragraphFour.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun paragraphFourRunOne = paragraphFour.createRun();
        paragraphFourRunOne.setBold(false);
        //paragraphFourRunOne.setUnderline(UnderlinePatterns.SINGLE);
        paragraphFourRunOne.setFontSize(10);
        paragraphFourRunOne.setFontFamily("Verdana");
        paragraphFourRunOne.setColor("000000");
        paragraphFourRunOne.setText(texto);
        
        FileOutputStream outStream = null;
        outStream = new FileOutputStream(caminhoDoc);
        
        documento.write(outStream);
        outStream.close();
	}
	
	/**
	 * Insere uma imagem no Documento Word.
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param caminhoImagem - String
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws AWTException
	 */
	public static void inserirImagem(String caminhoDoc, CustomXWPFDocument documento, String caminhoImagem) throws IOException, InvalidFormatException, AWTException{
		//Tira um print da tela
		Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
		BufferedImage capture = new Robot().createScreenCapture(screenRect);
		ImageIO.write(capture, "png", new File(caminhoImagem));
		
		//Insire ela no documento do word
		XWPFParagraph paragraphX = documento.getLastParagraph();
		paragraphX.setAlignment(ParagraphAlignment.LEFT);
		String blipId = paragraphX.getDocument().addPictureData(new FileInputStream(new File(caminhoImagem)),
		        Document.PICTURE_TYPE_JPEG);
		documento.createPicture(blipId,documento.getNextPicNameNumber(Document.PICTURE_TYPE_PNG),260, 145);

        FileOutputStream outStream = null;
    
        
        outStream = new FileOutputStream(caminhoDoc);
        documento.write(outStream);
        outStream.close();
	}
}
