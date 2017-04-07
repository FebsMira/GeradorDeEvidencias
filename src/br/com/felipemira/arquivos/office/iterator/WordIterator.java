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
import java.math.BigInteger;

import javax.imageio.ImageIO;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import br.com.felipemira.arquivos.office.custom.document.CustomXWPFDocument;

public class WordIterator {

	/**
	 * Insere um Titulo no arquivo word
	 * 
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param titulo - String
	 * @throws IOException
	 */
	public static void inserirTitulo(String caminhoDoc, CustomXWPFDocument documento, String titulo)
			throws IOException {
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
	 * 
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param texto - String
	 * @throws IOException
	 */
	public static void inserirTexto(String caminhoDoc, CustomXWPFDocument documento, String texto) throws IOException {
		XWPFParagraph paragraphFour = documento.createParagraph();
		paragraphFour.setAlignment(ParagraphAlignment.LEFT);
		XWPFRun paragraphFourRunOne = paragraphFour.createRun();
		paragraphFourRunOne.setBold(false);
		// paragraphFourRunOne.setUnderline(UnderlinePatterns.SINGLE);
		paragraphFourRunOne.setFontSize(9);
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
	 * 
	 * @param caminhoDoc - String
	 * @param documento - CustomXWPFDocument
	 * @param caminhoImagem - String
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws AWTException
	 */
	public static void inserirImagem(String caminhoDoc, CustomXWPFDocument documento, String caminhoImagem)
			throws IOException, InvalidFormatException, AWTException {
		// Tira um print da tela
		Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
		BufferedImage capture = new Robot().createScreenCapture(screenRect);
		ImageIO.write(capture, "png", new File(caminhoImagem));

		// Insire ela no documento do word
		XWPFParagraph paragraphX = documento.getLastParagraph();
		paragraphX.setAlignment(ParagraphAlignment.LEFT);
		String blipId = paragraphX.getDocument().addPictureData(new FileInputStream(new File(caminhoImagem)),
				Document.PICTURE_TYPE_JPEG);
		documento.createPicture(blipId, documento.getNextPicNameNumber(Document.PICTURE_TYPE_PNG), 260, 145);

		FileOutputStream outStream = null;

		outStream = new FileOutputStream(caminhoDoc);
		documento.write(outStream);
		outStream.close();
	}

	/**
	 * Insere dados em uma tabela especificada pelo numero.
	 * 
	 * @param caminhoDoc - String com o caminho.
	 * @param documento - CustomXWPFDocument.
	 * @param numeroTabela - int - Comeca com 0.
	 * @param numeroLinha - int - comeca com 0.
	 * @param numeroColuna - int - comeca com 0.
	 * @param dado - String a ser inserida.
	 * @throws IOException
	 */
	public static void inserirDadoTabela(String caminhoDoc, CustomXWPFDocument documento, int numeroTabela,
			int numeroLinha, int numeroColuna, String dado, boolean negrito) throws IOException {
		XWPFTable table = documento.getTableArray(numeroTabela);

		XWPFRun paragrafo = table.getRow(numeroLinha).getCell(numeroColuna).getParagraphs().get(0).createRun();
		paragrafo.getText(0);
		paragrafo.setBold(negrito);
		paragrafo.setFontSize(9);
		paragrafo.setFontFamily("Verdana");
		paragrafo.setColor("000000");
		paragrafo.setText(" " + dado);

		FileOutputStream outStream = null;
		outStream = new FileOutputStream(caminhoDoc);

		documento.write(outStream);
		outStream.close();

	}

	/**
	 * Insere dados em uma tabela especificada pelo numero.
	 * 
	 * @param caminhoDoc - String com o caminho.
	 * @param documento - CustomXWPFDocument.
	 * @param numeroTabela - int - Comeca com 0.
	 * @param numeroLinha - int - comeca com 0.
	 * @param numeroColuna - int - comeca com 0.
	 * @param dado - String a ser inserida.
	 * @throws IOException
	 */
	public static void inserirDadoTabelaLetraBranca(String caminhoDoc, CustomXWPFDocument documento, int numeroTabela,
			int numeroLinha, int numeroColuna, String dado, boolean negrito) throws IOException {
		XWPFTable table = documento.getTableArray(numeroTabela);

		XWPFRun paragrafo = table.getRow(numeroLinha).getCell(numeroColuna).getParagraphs().get(0).createRun();
		paragrafo.getText(0);
		paragrafo.setBold(negrito);
		paragrafo.setFontSize(9);
		paragrafo.setFontFamily("Verdana");
		paragrafo.setColor("ffffff");
		paragrafo.setText(" " + dado);

		FileOutputStream outStream = null;
		outStream = new FileOutputStream(caminhoDoc);

		documento.write(outStream);
		outStream.close();

	}

	public static void newPage(String caminhoDoc, CustomXWPFDocument document) {
		XWPFParagraph paragraph = document.createParagraph();
		paragraph.setPageBreak(true);

	}

	public static void newTableFuncional(String caminhoDoc, CustomXWPFDocument document, int numeroIterator)
			throws Exception {
		// Cria uma nova tabela
		XWPFTable table = document.createTable(7, 2);

		// Faz um merge entre celulas
		table.getRow(5).getCell(0).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
		table.getRow(5).getCell(1).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);

		int tamanho = 1440;

		// Define a altura da linha
		table.getRow(5).setHeight((int) (tamanho * 2 / 10)); // set height 1/10
																// inch.

		// Define a altura da linha
		table.getRow(6).setHeight((int) (tamanho * 2 / 10)); // set height 1/10
																// inch.

		// Cor de fundo da linha
		table.getRow(5).getCell(0).setColor("ff8000");

		// Alinha o texto dentro do paragrafo
		table.getRow(5).getCell(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
		table.getRow(5).getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);

		// Faz um merge entre celulas
		table.getRow(6).getCell(0).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
		table.getRow(6).getCell(1).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);

		// Cor de fundo da linha
		table.getRow(6).getCell(0).setColor("f2f2f2");

		// Tira a borda dos lados em determinada celula
		table.getRow(4).getCell(0).getCTTc().addNewTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.NIL);
		table.getRow(4).getCell(0).getCTTc().addNewTcPr().addNewTcBorders().addNewRight().setVal(STBorder.NIL);
		table.getRow(4).getCell(1).getCTTc().addNewTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.NIL);
		table.getRow(4).getCell(1).getCTTc().addNewTcPr().addNewTcBorders().addNewRight().setVal(STBorder.NIL);

		inserirDadoTabela(caminhoDoc, document, numeroIterator, 0, 0, "Passo", true);
		inserirDadoTabela(caminhoDoc, document, numeroIterator, 1, 0, "Descrição:", true);
		inserirDadoTabela(caminhoDoc, document, numeroIterator, 2, 0, "Resultado Esperado:", true);
		inserirDadoTabela(caminhoDoc, document, numeroIterator, 3, 0, "Resultado Obtido:", true);
		inserirDadoTabela(caminhoDoc, document, numeroIterator, 3, 1, "-", true);

		inserirDadoTabelaLetraBranca(caminhoDoc, document, numeroIterator, 5, 0, "Validação", true);

		for (int i = 0; i < 4; i++) {
			// Define a largura das colunas
			table.getRow(i).getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(tamanho * 3));
			table.getRow(i).getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(tamanho * 7));
			// Alinha o texto dentro da linha
			table.getRow(i).getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);
			// Define cor da linha
			table.getRow(i).getCell(0).setColor("f2f2f2");
		}

	}
}
