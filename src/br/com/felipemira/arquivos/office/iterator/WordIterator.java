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
	
	/**
	 * Insere dados em uma tabela especificada pelo número.
	 * @param caminhoDoc - String com o caminho.
	 * @param documento - CustomXWPFDocument.
	 * @param numeroTabela - int - Começa com 0.
	 * @param numeroLinha - int - começa com 0.
	 * @param numeroColuna - int - começa com 0.
	 * @param dado - String a ser inserida.
	 * @throws IOException
	 */
	public static void inserirDadoTabela(String caminhoDoc, CustomXWPFDocument documento, int numeroTabela, int numeroLinha, int numeroColuna, String dado) throws IOException{
		XWPFTable table = documento.getTableArray(numeroTabela);
		
		XWPFRun casoDeTeste = table.getRow(numeroLinha).getCell(numeroColuna).getParagraphs().get(0).createRun();
		casoDeTeste.setBold(false);
		casoDeTeste.setFontSize(10);
		casoDeTeste.setFontFamily("Verdana");
		casoDeTeste.setColor("000000");
		casoDeTeste.setText(" " + dado);
		
		FileOutputStream outStream = null;
        outStream = new FileOutputStream(caminhoDoc);
        
        documento.write(outStream);
        outStream.close();
		
	}
}
