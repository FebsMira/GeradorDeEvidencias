package br.com.felipemira.view;

import java.io.File;
import java.io.IOException;

import br.com.felipemira.RftToWord;
import br.com.felipemira.objects.object.Error;
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
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseButton;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.AnchorPane;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;

public class GeradorDeEvidencias extends Application{
	
	private Scene scene;
	private Image applicationIcon;
	
	private AnchorPane pane;
	private Label labelSalvar, labelArquivoRTF, labelLinhaInicio, labelLinhaFim, labelBy; //labelModel
	private TextField txSalvar, txArquivoRTF, txLinhaInicio, txLinhaFinal; //txModel
	private Button btnExecutar;
	private ProgressBar progressBar;
	private ImageView imagem;
	
	private String localIconeAplicativo = "br/com/felipemira/img/Inmetrics.png";
	
	
	private static Stage stage;
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private Service<Void> servico = new Service(){
		@Override
        protected Task createTask() {
            return new Task() {
            	@Override
                protected Void call() throws NumberFormatException, InterruptedException, IOException {
            			RftToWord.executar(txSalvar.getText(), "templates"/*txModel.getText()*/, txArquivoRTF.getText(), Integer.valueOf(txLinhaInicio.getText()), Integer.valueOf(txLinhaFinal.getText()));
            		return null;
            	}
        	};
	 	}	
	};
	
	ChangeListener<Worker.State> listener = new ChangeListener<Worker.State>() {
        @Override
        public void changed(ObservableValue<? extends Worker.State> observableValue, Worker.State oldState, Worker.State newState) {
            if (newState == Worker.State.SUCCEEDED) {
            	progressBar.progressProperty().unbind();
            	progressBar.setProgress(new Float(0f));
            	alert("Finalizado!", "Criação de documentos referentes a RTF finalizada!");
            }
            if(newState == Worker.State.FAILED){
            	progressBar.progressProperty().unbind();
            	progressBar.setProgress(new Float(0f));
            	alertError("Erro!", Error.error);
            }
        }
	};
	
	@Override
	public void start(Stage stage) throws Exception {
		initComponents();
		initListeners();
		
		stage.setScene(scene);
		stage.setTitle("Gerador de Evidências");
		stage.resizableProperty().set(false);
		applicationIcon = new Image(localIconeAplicativo);
		stage.getIcons().add(applicationIcon);
		stage.show();
		
		initLayout();
		GeradorDeEvidencias.stage = stage;
	}
	
	public static void main(String[] args){
		launch(args);
	}
	
	public static Stage getStage(){
		return stage;
	}
	
	private void initComponents(){
		pane = new AnchorPane();
		pane.setPrefSize(800, 300);
		pane.setStyle("-fx-background-color:linear-gradient(from 200% 0% to 100% 100%, white 0%, #FFD401 100%)");
		
		
		/*labelModel = new Label("Local do Modelo.docx:");
		txModel = new TextField();
		txModel.setPromptText("Local do Modelo.docx.");*/
		
		labelSalvar = new Label("Salvar em:");
		txSalvar = new TextField();
		txSalvar.setPromptText("Salvar em...");
		
		labelArquivoRTF = new Label("Arquivo RTF:");
		txArquivoRTF = new TextField();
		txArquivoRTF.setPromptText("Arquivo RTF...");
		
		labelLinhaInicio = new Label("Linha Início:");
		txLinhaInicio = new TextField();
		txLinhaInicio.setPromptText(">=12");
		txLinhaInicio.textProperty().addListener(new ChangeListener<String>() {
	        @Override
	        public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
	            if (!newValue.matches("\\d*")) {
	            	txLinhaInicio.setText(newValue.replaceAll("[^\\d]", ""));
	            }
	        }
	    });
		
		labelLinhaFim = new Label("Linha Fim:");
		txLinhaFinal = new TextField();
		txLinhaFinal.setPromptText(">=12");
		txLinhaFinal.textProperty().addListener(new ChangeListener<String>() {
	        @Override
	        public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
	            if (!newValue.matches("\\d*")) {
	            	txLinhaFinal.setText(newValue.replaceAll("[^\\d]", ""));
	            }
	        }
	    });
		
		btnExecutar = new Button("Executar");
		
		imagem = new ImageView();
		imagem.setImage(new Image("br/com/felipemira/img/Inmetrics2.jpg"));
		
		labelBy = new Label("Criado por Felipe Mira");
		
		progressBar = new ProgressBar();
		progressBar.setProgress(new Float(0f));
		
		pane.getChildren().addAll(labelSalvar, txSalvar, labelArquivoRTF, txArquivoRTF, labelLinhaInicio, txLinhaInicio, labelLinhaFim, txLinhaFinal, btnExecutar, progressBar, imagem, labelBy);//, labelBy, labelModel, txModel);
		scene = new Scene(pane);
	}
	
	private void initLayout(){
		/*
		labelModel.setLayoutX(((pane.getWidth() - labelModel.getWidth()) / 2) - 300);
		labelModel.setLayoutY(20);
		txModel.setLayoutX(((pane.getWidth() - txModel.getWidth()) / 2) - 100);
		txModel.setPrefWidth(500);
		txModel.setLayoutY(20);*/
		
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
		
	}
	
	private void initListeners(){
		/*
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
		});*/
		
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
		    	/*if(txModel.getText().equals("")){
		    		mensagem = mensagem + "- Local do Modelo.docx\n";
		    		erro = true;
		    	}*/
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
		    		alert("Atenção!", mensagem);
		    	}else{
		    		String mensagem2 = "Verificar valor(es):\n";
		    		if(Integer.valueOf(Integer.parseInt(txLinhaInicio.getText())) < 12){
		    			mensagem2 = mensagem2 + "- Linha Início Menor que 12.\n";
			    		erro = true;
		    		}
		    		if(Integer.valueOf(Integer.parseInt(txLinhaFinal.getText())) < 12){
		    			mensagem2 = mensagem2 + "- Linha Fim Menor que 12.\n";
			    		erro = true;
		    		}
		    		if(erro){
		    			alert("Atenção!", mensagem);
		    		}else{
		    			try {
			    			progressBar.progressProperty().bind(servico.progressProperty());
			    			if(servico.getState() == State.SUCCEEDED || servico.getState() == State.FAILED){
			    				servico.reset();
			    				servico.stateProperty().removeListener(listener);
			    				servico.stateProperty().addListener(listener);
			    				servico.start();
			    			}else{
			    				servico.stateProperty().removeListener(listener);
			    				servico.stateProperty().addListener(listener);
			    				servico.start();
			    			}
						} catch (NumberFormatException e1) {
							e1.printStackTrace();
						}
		    		}
		    	}
		    }
		});
	}
	
	private String selecionarPasta(String texto){
		DirectoryChooser directoryChooser = new DirectoryChooser();
		directoryChooser.setInitialDirectory(new java.io.File("."));
		directoryChooser.setTitle(texto);
		File selectedDirectory = directoryChooser.showDialog(stage);
		if(selectedDirectory == null){
			return "null";
		}else{
			return selectedDirectory.toString();
		}
	}
	
	private String selecionarArquivo(String texto){
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new java.io.File("."));
	    fileChooser.setTitle(texto);
	    fileChooser.getExtensionFilters().addAll(
	            new ExtensionFilter("Excel Files", "*.xlsx"));
	    File file = fileChooser.showOpenDialog(stage);

	    if (file == null){
	        return "null";
	    }else{
	    	return file.getPath();
	    }
	}
	
	private void alert(String title, String text){
		Alert alert = new Alert(AlertType.INFORMATION);
		
		alert.setTitle(title);
    	alert.setHeaderText(text);
    	alert.setContentText(stage.getTitle().toString());
    	
		Stage stage = (Stage) alert.getDialogPane().getScene().getWindow();
		stage.getIcons().add(new Image(localIconeAplicativo));
    	
    	alert.showAndWait().ifPresent(rs -> {
    	    if (rs == ButtonType.OK) {
    	        System.out.println("Pressed OK.");
    	    }
    	});
	}
	
	private void alertError(String title, String text){
		Alert alert = new Alert(AlertType.ERROR);
		
		alert.setTitle(title);
    	alert.setHeaderText(text);
    	alert.setContentText(stage.getTitle().toString());
    	
		Stage stage = (Stage) alert.getDialogPane().getScene().getWindow();
		stage.getIcons().add(new Image(localIconeAplicativo));
    	
    	alert.showAndWait().ifPresent(rs -> {
    	    if (rs == ButtonType.OK) {
    	        System.out.println("Pressed OK.");
    	    }
    	});
	}
}
