package br.com.felipemira.view;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import br.com.felipemira.RftToWord;
import javafx.application.Application;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.concurrent.Service;
import javafx.concurrent.Task;
import javafx.concurrent.Worker;
import javafx.concurrent.Worker.State;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseButton;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.AnchorPane;
import javafx.stage.Stage;

public class GeradorDeEvidencias extends Application{
	
	@Override
	public void start(Stage stage) throws Exception {
		AnchorPane pane = new AnchorPane();
		pane.setPrefSize(800, 300);
		pane.setStyle("-fx-background-color:linear-gradient(from 200% 0% to 100% 100%, white 0%, #FFD401 100%)");
		
		Label labelModel = new Label("Local do Modelo.docx:");
		TextField txModel = new TextField();
		txModel.setPromptText("Local do Modelo.docx.");
		
		Label labelSalvar = new Label("Salvar em:");
		TextField txSalvar = new TextField();
		txSalvar.setPromptText("Salvar em...");
		
		Label labelArquivoRTF = new Label("Arquivo RTF:");
		TextField txArquivoRTF = new TextField();
		txArquivoRTF.setPromptText("Arquivo RTF...");
		
		Label labelLinhaInicio = new Label("Linha Início:");
		TextField txLinhaInicio = new TextField();
		txLinhaInicio.setPromptText(">=12");
		txLinhaInicio.textProperty().addListener(new ChangeListener<String>() {
	        @Override
	        public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
	            if (!newValue.matches("\\d*")) {
	            	txLinhaInicio.setText(newValue.replaceAll("[^\\d]", ""));
	            }
	        }
	    });
		
		
		Label labelLinhaFim = new Label("Linha Fim:");
		TextField txLinhaFinal = new TextField();
		txLinhaFinal.setPromptText(">=12");
		txLinhaFinal.textProperty().addListener(new ChangeListener<String>() {
	        @Override
	        public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
	            if (!newValue.matches("\\d*")) {
	            	txLinhaFinal.setText(newValue.replaceAll("[^\\d]", ""));
	            }
	        }
	    });
		
		Button btnExecutar = new Button("Executar");
		
		ProgressBar progressBar = new ProgressBar();
		
		@SuppressWarnings({ "unchecked", "rawtypes" })
		Service<Void> servico = new Service(){
		 @SuppressWarnings("rawtypes")
		@Override
            protected Task createTask() {
                return new Task() {
                	@Override
                    protected Void call() throws Exception {
                		RftToWord.executar(txSalvar.getText(), txModel.getText(), txArquivoRTF.getText(), Integer.valueOf(txLinhaInicio.getText()), Integer.valueOf(txLinhaFinal.getText()));
                		return null;
                	}
            	};
		 	}	
		};
		progressBar.setProgress(new Float(0f));
	    
		ImageView imagem = new ImageView();
		imagem.setImage(new Image(getClass().getResourceAsStream("Inmetrics2.jpg")));
		
		Label labelBy = new Label("By Felipe Mira");
		
		pane.getChildren().addAll(labelModel, txModel, labelSalvar, txSalvar, labelArquivoRTF, txArquivoRTF, labelLinhaInicio, txLinhaInicio, labelLinhaFim, txLinhaFinal, btnExecutar, progressBar, imagem, labelBy);
		
		Scene scene = new Scene(pane);
		stage.setScene(scene);
		stage.setTitle("Gerador de Evidências");
		stage.resizableProperty().set(false);
		
		Image applicationIcon = new Image(getClass().getResourceAsStream("Inmetrics.png"));
		stage.getIcons().add(applicationIcon);
		
		stage.show();
		
		
		labelModel.setLayoutX(((pane.getWidth() - labelModel.getWidth()) / 2) - 300);
		labelModel.setLayoutY(20);
		txModel.setLayoutX(((pane.getWidth() - txModel.getWidth()) / 2) - 100);
		txModel.setPrefWidth(500);
		txModel.setLayoutY(20);
		
		labelSalvar.setLayoutX(((pane.getWidth() - labelSalvar.getWidth()) / 2) - 332);
		labelSalvar.setLayoutY(60);
		txSalvar.setLayoutX(((pane.getWidth() - txSalvar.getWidth()) / 2) - 100);
		txSalvar.setPrefWidth(500);
		txSalvar.setLayoutY(60);
		
		labelArquivoRTF.setLayoutX(((pane.getWidth() - labelArquivoRTF.getWidth()) / 2) - 325);
		labelArquivoRTF.setLayoutY(100);
		txArquivoRTF.setLayoutX(((pane.getWidth() - txArquivoRTF.getWidth()) / 2) - 100);
		txArquivoRTF.setPrefWidth(500);
		txArquivoRTF.setLayoutY(100);
		
		labelLinhaInicio.setLayoutX(((pane.getWidth() - labelLinhaInicio.getWidth()) / 2) - 145);
		labelLinhaInicio.setLayoutY(140);
		txLinhaInicio.setLayoutX(((pane.getWidth() - txArquivoRTF.getWidth()) / 2) - 10);
		txLinhaInicio.setPrefWidth(50);
		txLinhaInicio.setLayoutY(140);
		
		labelLinhaFim.setLayoutX(((pane.getWidth() - labelLinhaFim.getWidth()) / 2) + 220);
		labelLinhaFim.setLayoutY(140);
		txLinhaFinal.setLayoutX(((pane.getWidth() - txLinhaFinal.getWidth()) / 2) + 350);
		txLinhaFinal.setPrefWidth(50);
		txLinhaFinal.setLayoutY(140);
		
		btnExecutar.setLayoutX(((pane.getWidth() - btnExecutar.getWidth()) / 2) + 80);
		btnExecutar.setLayoutY(210);
		
		imagem.setLayoutX(((pane.getWidth() - btnExecutar.getWidth()) / 2) - 330);
		imagem.setFitHeight(70);
		imagem.setFitWidth(170);
		imagem.setLayoutY(190);
		
		progressBar.setLayoutX(((pane.getWidth() - btnExecutar.getWidth()) / 2) + 260);
		progressBar.setLayoutY(280);
		
		labelBy.setLayoutX(((pane.getWidth() - btnExecutar.getWidth()) / 2) - 320);
		labelBy.setScaleY(0.5);
		labelBy.setScaleX(0.5);
		labelBy.setLayoutY(280);
		
		txModel.setOnMouseClicked(new EventHandler<MouseEvent>() {
		    @Override
		    public void handle(MouseEvent mouseEvent) {
		        if(mouseEvent.getButton().equals(MouseButton.PRIMARY)){
		            if(mouseEvent.getClickCount() == 1){
		            	String retorno = selecionarPasta("Local do Modelo.docx");
			        	txModel.setText((retorno.equals("null"))? "" : retorno);
		            }
		        }
		    }
		});
		
		txSalvar.setOnMouseClicked(new EventHandler<MouseEvent>() {
		    @Override
		    public void handle(MouseEvent mouseEvent) {
		        if(mouseEvent.getButton().equals(MouseButton.PRIMARY)){
		            if(mouseEvent.getClickCount() == 1){
		            	String retorno = selecionarPasta("Salvar em:");
			        	txSalvar.setText((retorno.equals("null"))? "" : retorno);
		            }
		        }
		    }
		});
		
		txArquivoRTF.setOnMouseClicked(new EventHandler<MouseEvent>() {
		    @Override
		    public void handle(MouseEvent mouseEvent) {
		        if(mouseEvent.getButton().equals(MouseButton.PRIMARY)){
		            if(mouseEvent.getClickCount() == 1){
		            	String retorno = selecionarArquivo("Arquivo RTF:");
			        	txArquivoRTF.setText((retorno.equals("null"))? "" : retorno);
		            }
		        }
		    }
		});
		
		btnExecutar.setOnAction(new EventHandler<ActionEvent>() {
		    @Override public void handle(ActionEvent e) {
		    	Boolean erro = false;
		    	String mensagem = "Preencha o(s) seguinte(s) campo(s):\n";
		    	if(txModel.getText().equals("")){
		    		mensagem = mensagem + "- Local do Modelo.docx\n";
		    		erro = true;
		    	}
		    	if(txSalvar.getText().equals("")){
		    		mensagem = mensagem + "- Salvar em:\n";
		    		erro = true;
		    	}
		    	if(txArquivoRTF.getText().equals("")){
		    		mensagem = mensagem + "- Arquivo RTF:\n";
		    		erro = true;
		    	}
		    	if(txLinhaInicio.getText().equals("")){
		    		mensagem = mensagem + "- Linha Início:\n";
		    		erro = true;
		    	}
		    	if(txLinhaFinal.getText().equals("")){
		    		mensagem = mensagem + "- Linha Fim:\n";
		    		erro = true;
		    	}
		    	if(erro){
		    		JOptionPane.showMessageDialog(null, mensagem, "Atenção!", JOptionPane.INFORMATION_MESSAGE);
		    	}else{
		    		String mensagem2 = "Verificar valor(es):\n";
		    		if(Integer.valueOf(txLinhaInicio.getText()) < 12){
		    			mensagem2 = mensagem2 + "- Linha Início Menor que 12.\n";
			    		erro = true;
		    		}
		    		if(Integer.valueOf(txLinhaFinal.getText()) < 12){
		    			mensagem2 = mensagem2 + "- Linha Fim Menor que 12.\n";
			    		erro = true;
		    		}
		    		if(erro){
		    			JOptionPane.showMessageDialog(null, mensagem2, "Atenção!", JOptionPane.INFORMATION_MESSAGE);
		    		}else{
		    			try {
			    			progressBar.progressProperty().bind(servico.progressProperty());
			    			if(servico.getState() == State.SUCCEEDED){
			    				servico.reset();
			    				servico.stateProperty().addListener(new ChangeListener<Worker.State>() {
			                        @Override
			                        public void changed(ObservableValue<? extends Worker.State> observableValue, Worker.State oldState, Worker.State newState) {
			                            if (newState == Worker.State.SUCCEEDED) {
			                            	progressBar.progressProperty().unbind();
			                            	progressBar.setProgress(new Float(0f));
			                            	JOptionPane.showMessageDialog(null, "Criação de documentos referentes a RTF finalizado!", "Finalizado!", JOptionPane.INFORMATION_MESSAGE);
			                            }
			                        }
			                    });
			    				servico.start();
			    			}else{
			    				servico.stateProperty().addListener(new ChangeListener<Worker.State>() {
			                        @Override
			                        public void changed(ObservableValue<? extends Worker.State> observableValue, Worker.State oldState, Worker.State newState) {
			                            if (newState == Worker.State.SUCCEEDED) {
			                            	progressBar.progressProperty().unbind();
			                            	progressBar.setProgress(new Float(0f));
			                            	JOptionPane.showMessageDialog(null, "Criação de documentos referentes a RTF finalizado!", "Finalizado!", JOptionPane.INFORMATION_MESSAGE);
			                            }
			                        }
			                    });
			    				servico.start();
			    			}
						} catch (NumberFormatException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
		    		}
		    	}
		    }
		});
	}
	
	public static void main(String[] args){
		launch(args);
	}
	
	public String selecionarPasta(String texto){
		JFileChooser chooser = new JFileChooser();
		chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle(texto);
		chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		chooser.setAcceptAllFileFilterUsed(false);

		if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
		  System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
		  System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
		} else {
		  System.out.println("No Selection ");
		}
		return String.valueOf(chooser.getSelectedFile());
	}
	
	public String selecionarArquivo(String texto){
		JFileChooser chooser = new JFileChooser();
		chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle(texto);
		chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
		FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", "xlsx");
		chooser.setFileFilter(filter);
		chooser.setAcceptAllFileFilterUsed(false);

		if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
		  System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
		  System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
		} else {
		  System.out.println("No Selection ");
		}
		return String.valueOf(chooser.getSelectedFile());
	}
}
