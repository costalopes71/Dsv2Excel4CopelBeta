package br.com.sinapsis.exportadorcopel.main;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;

import br.com.sinapsis.exportadorcopel.exporter.ExportadorCopel;

public class Main {

	public static void main(String[] args) throws ParseException {
		
		String originFilePath = "C:/Users/usrsnp/Documents/Joao/ArquivoModeloCopel/LEITURA_ALIMENTADOR.dsv";
		String destinyFilePath = "C:/Users/usrsnp/Documents/Joao/ArquivoModeloCopel/processed_excel_file_TESTE.xlsx";
		String referencePath = "C:/Users/usrsnp/Documents/Joao/ArquivoModeloCopel/SubAlim.csv";
		
		try {
			ExportadorCopel e = new ExportadorCopel(originFilePath, destinyFilePath, referencePath);
			e.generateExcelFile(true);
//			e.orderDsvFileBySubstation();
//			e.generateExcelTemplate(true);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		
	}
	
}
