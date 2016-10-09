package br.com.felipemira.copy;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class Copy {
	public static void copyFile(File origem, File destino) throws IOException{
		InputStream in = new FileInputStream(origem);
		OutputStream out = new FileOutputStream(destino);
		
		byte[] buffer = new byte[1024];
		int lenght;
		while((lenght = in.read(buffer)) > 0){
			out.write(buffer, 0, lenght);
		}
		in.close();
		out.close();
	}
}
