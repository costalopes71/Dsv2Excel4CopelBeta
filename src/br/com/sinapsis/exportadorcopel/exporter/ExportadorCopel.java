package br.com.sinapsis.exportadorcopel.exporter;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import br.com.sinapsis.exportadorcopel.entities.Alimentador;
import br.com.sinapsis.exportadorcopel.entities.Medicao;
import br.com.sinapsis.exportadorcopel.entities.Subestacao;
import br.com.sinapsis.exportadorcopel.util.Utils;

/**
 * Classe responsável por exportar os dados de um arquivo DSV da empresa Copel que contém as medições de seus alimentadores
 * e criar um arquivo Excel (xlsx) com um determinado template inserindo essas medições. Esse arquivo excel é usado
 * posteriormente pelo programa ConversorDeMedicoes para gerar o arquivo MDB que posteriormente irá popular um banco de dados.
 * @author Joao Lopes
 * @since 10/01/2018
 * @version 1.0b
 */
public class ExportadorCopel {

	private SXSSFWorkbook wb;
	private Utils util;
	private BufferedReader br;
	private String dsvPath;
	private String excelPath;
	private String referenceFilePath;
	private long startTime;
	private HashMap<String, Integer> mapPosicaoAlimentores;
	private HashMap<String, HashMap<String, Integer>> posicoesLinha;

	public ExportadorCopel(String dsvPath, String excelPath, String referenceFilePath) throws FileNotFoundException {
		wb = new SXSSFWorkbook(-1);
		this.dsvPath = dsvPath;
		this.excelPath = excelPath;
		this.referenceFilePath = referenceFilePath;
		util = new Utils();
		startTime = System.currentTimeMillis();
		mapPosicaoAlimentores = new HashMap<>();
		posicoesLinha = new HashMap<>();
	}

	/**
	 * Método que gera um arquivo excel para a empresa Copel a partir de um arquivo DSV.
	 * @param hasHeader <code>boolean</code>, booleano que indica se o arquivo DSV possui cabeçalho.
	 * @throws IOException
	 * @throws ParseException
	 */
	public void generateExcelFile(boolean hasHeader) throws IOException, ParseException {
		
		this.orderDsvFileBySubstation();
		this.generateExcelTemplate(hasHeader);
		
		long start = System.currentTimeMillis();
		System.out.println("----------------------------------------------------------------");
		System.out.println("Initiating the data write on the file.");
		System.out.println("----------------------------------------------------------------");
		System.out.println("\nWritting data, please wait. This can take several minutes depending on the size of the file\n");

		br = util.openBufferedReader(dsvPath);
		util.hasHeader(true, br);
		
		String line;
		String lastSubstation = "";
		while ((line = br.readLine()) != null) {
			String[] fields = line.split("\\|");
			
			/*esta condicao serve para quando mudar a subestacao (uma vez que o arquivo DSV foi ordenado por subestacao pelo
			método orderDsvFileBySubstation() ) fazer o flush daquela subestacao (ou seja, daquela sheet uma vez que cada
			sheet é um subestacao), isto é feito para evitar o estouro de memória (é feito a cada subestacao pq uma vez que
			as linhas sao flushadas elas não estão mais acessíveis em memória).*/
			if (!lastSubstation.equals(fields[3]) && !lastSubstation.equals("")) {
				SXSSFSheet sheetToFlush = wb.getSheet(lastSubstation);
				sheetToFlush.flushRows();
			}
			
			String siglaSub = fields[3];
			Medicao medicao = this.setMedicaoFromFields(fields);
			SXSSFSheet sheet = wb.getSheet(fields[3]);
			
			int initialCell = mapPosicaoAlimentores.get(fields[0]);
			int rowNum = 0;
			
			if (posicoesLinha.containsKey(siglaSub)) {
				if (posicoesLinha.get(siglaSub).containsKey(fields[2])) {
					rowNum = posicoesLinha.get(siglaSub).get(fields[2]);
					Row row = sheet.getRow(rowNum);
					this.writeMedicao(row, medicao, initialCell);
				} else {
					Row newRow = this.writeRow(sheet, medicao);
					rowNum = newRow.getRowNum();
					this.writeMedicao(newRow, medicao, initialCell);
					posicoesLinha.get(siglaSub).put(fields[2], rowNum);
				}
			} else {
				Row newRow = this.writeRow(sheet, medicao);
				rowNum = newRow.getRowNum();
				this.writeMedicao(newRow, medicao, initialCell);
				HashMap<String, Integer> aux = new HashMap<>();
				aux.put(fields[2], rowNum);
				posicoesLinha.put(siglaSub, aux);
			}
			
			lastSubstation = fields[3];
		}
		
		util.save(wb, excelPath);
		br.close();
		
		double endTime = System.currentTimeMillis();
		double elapsed = endTime - start;
		double totalElapsedTime = endTime - startTime;
		
		System.out.println("----------------------------------------------------------------");
		System.out.println("SUCCESS: data written successfully .");
		System.out.println("Elapsed time: " + (elapsed / 1000)+ " seconds.");
		System.out.println("----------------------------------------------------------------\n");
		
		System.out.println("----------------------------------------------------------------");
		System.out.println("Thank you for using ExportadorCopel.");
		System.out.println("Elapsed TOTAL time: " + (totalElapsedTime / 1000)+ " seconds.");
		System.out.println("----------------------------------------------------------------\n");
	}
	
	private Row writeRow(SXSSFSheet sheet, Medicao medicao) {
		Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
		newRow.createCell(0).setCellValue(util.formatDatePattern(medicao.getData()));
		newRow.createCell(1).setCellValue(util.formatHourPattern(medicao.getData()));
		return newRow;
	}

	/**
	 * Método que gera o template excel onde os dados deverão posteriormente serem inseridos. Para tanto, este método
	 * lê o arquivo DSV da empresa Copel e para cada Subestação cria uma planilha (aba) no arquivo excel. Cria também
	 * o cabeçalho de cada planilha onde estarão a data, a hora da medição e os alimentadores (com suas respectivas
	 * medições) daquela Subestacao.
	 * @param hasHeader
	 * @throws IOException
	 */
	public void generateExcelTemplate(boolean hasHeader) throws IOException {

		long start = System.currentTimeMillis();
		System.out.println("----------------------------------------------------------------");
		System.out.println("Initiating the generation of the TEMPLATE file.");
		System.out.println("----------------------------------------------------------------");

		br = util.openBufferedReader(dsvPath);
		HashMap<String, Subestacao> subMap = new HashMap<>();
		util.hasHeader(hasHeader, br);

		String line;
		while ((line = br.readLine()) != null) {
			String[] fields = line.split("\\|");
			 /* Testa se no mapa ja existe a subestacao do registro que acabou de ser lido. Se não tiver ele cria um 
			  * novo objeto subestacao, seta o seu atributo codigo (codigo da subestacao no arquivo é o índice 3 do 
			  * vetor de campos), cria um objeto Alimentador e seta o seu codigo (codigo do alimentador no arquivo é 
			  * o índice 1 do vetor de campos), pega o atributo alimentadores (HashSet) do objeto subestacao que foi 
			  * criado e adiciona ao HashSet o alimentador. Finalmente, adiciona no HashMap subMap (linha 69 e 70) o 
			  * objeto subestacao. */
			if (!subMap.containsKey(fields[3])) {
				Subestacao subAux = new Subestacao();
				subAux.setSigla(fields[3]);
				Alimentador alimentador = new Alimentador();
				alimentador.setNome(fields[0]);
				subAux.getAlimentadores().add(alimentador);
				subMap.put(subAux.getSigla(), subAux);
				/* Se a subestacao ja existir testa se o alimentador daquele registro lido já existe dentro do HashSet da
				subestacao, se nao existir, adiciona. */
			} else {
				Subestacao sub = subMap.get(fields[3]);
				HashSet<Alimentador> alimentadorSet = sub.getAlimentadores();
				if (!alimentadorSet.contains(fields[0])) {
					Alimentador aux = new Alimentador();
					aux.setNome(fields[0]);
					alimentadorSet.add(aux);
					sub.setAlimentadores(alimentadorSet);
					subMap.put(sub.getSigla(), sub);
				}
			}
		}
		
		br.close();
		this.createSheets(subMap);
		double elapsed = System.currentTimeMillis() - start;

		System.out.println("----------------------------------------------------------------");
		System.out.println("SUCCESS: Template file generated.");
		System.out.println("Elapsed time: " + (elapsed / 1000) + " seconds.");
		System.out.println("----------------------------------------------------------------");
		
	}

	/**
	 * Método responsável por ordenar o arquivo DSV da empresa COPEL por subestação. Gera um novo arquivo DSV no mesmo
	 * caminho que o DSV original e com o mesmo nome mudando apenas o final com _ORDERED. Este novo arquivo será usado
	 * pelo método principal <code>generateExcelFile</code> para escrever os dados no arquivo excel sem causar problema
	 * de memória (out of memory).
	 * @throws IOException
	 */
	public void orderDsvFileBySubstation() throws IOException {
		
		long start = System.currentTimeMillis();
		System.out.println("----------------------------------------------------------------");
		System.out.println("Initiating the generation of the ordered by substation DSV file.");
		System.out.println("----------------------------------------------------------------");
		
		HashMap<Integer, String[]> referenceMap = this.createReferenceMap(referenceFilePath);
		br.close();
		
		br = util.openBufferedReader(dsvPath);
		String orderedDsvPath = this.createOrderedDsvPath();
		PrintWriter pw = util.createPrintWriter(orderedDsvPath);
		HashMap<String, ArrayList<String>> mapSub = new HashMap<>();
		String header = br.readLine();
		pw.println(header);
		
		String line = null;
		while ((line = br.readLine()) != null) {
			String[] fields = line.split("\\|");
			
			if (referenceMap.containsKey(Integer.parseInt(fields[0]))) {
				String[] referenceFields = referenceMap.get(Integer.parseInt(fields[0]));
				fields[0] = referenceFields[0]; //nome do alimentador
				fields[3] = referenceFields[1]; //sigla da subestação
			} else {
				//se nao contem manda para o LOG o numero da SUB e o numero do ALIM.
			}
			
			StringBuilder sb = new StringBuilder();
			
			for (int i = 0; i < fields.length; i++) {
				sb.append(fields[i]);
				if (i != fields.length - 1) {
					sb.append("|");
				}
			}
			
			String newLine = sb.toString();
			String siglaSub = fields[3];
			if (mapSub.containsKey(siglaSub)) {
				mapSub.get(siglaSub).add(newLine);
			} else {
				ArrayList<String> aux = new ArrayList<>();
				aux.add(newLine);
				mapSub.put(siglaSub, aux);
			}
		}
		
		for (Map.Entry<String, ArrayList<String>> mapValue : mapSub.entrySet()) {
			ArrayList<String> aux = mapValue.getValue();
			for (int i = 0; i < aux.size(); i++) {
				pw.println(aux.get(i));
			}
		}
		
		br.close();
		pw.close();
		this.dsvPath = orderedDsvPath;
		
		double elapsed = System.currentTimeMillis() - start;
		System.out.println("----------------------------------------------------------------");
		System.out.println("SUCCESS: Ordered file generated.");
		System.out.println("Elapsed time: " + (elapsed / 1000) + " seconds.");
		System.out.println("----------------------------------------------------------------");
		
	}
	
	private String createOrderedDsvPath() {
		return dsvPath.substring(0, dsvPath.lastIndexOf(".")) + "_ORDERED.dsv";
	}

	private void createSheets(HashMap<String, Subestacao> subMap) {
		for (Map.Entry<String, Subestacao> mapValue : subMap.entrySet()) {
			Subestacao subestacao = mapValue.getValue();
			SXSSFSheet sheet = wb.createSheet(subestacao.getSigla());
			Row rowHeader = sheet.createRow(0);
			createHeaderSheet(rowHeader, subestacao);
		}
	}
	
	private void createHeaderSheet(Row rowHeader, Subestacao subestacao) {
		this.createStaticHeaderFields(rowHeader);
		int cellCounter = 2;
		int qtAlim = subestacao.getAlimentadores().size();
		
		for (Alimentador alimentador : subestacao.getAlimentadores()) {
			
			if (!mapPosicaoAlimentores.containsKey(alimentador.getNome())) {
				mapPosicaoAlimentores.put(alimentador.getNome(), cellCounter);
			}
			
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_IA");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_IB");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_IC");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_POT_ATIVA");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_POT_REAT");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_FATOR_POT");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_TENSAOA");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_TENSAOB");
			cellCounter++;
			rowHeader.createCell(cellCounter).setCellValue(subestacao.getSigla() + "_" + alimentador.getNome() + "_TENSAOC");
			cellCounter++;
		}
		this.makeHeaderBold(rowHeader, qtAlim);
	}

	private void makeHeaderBold(Row rowHeader, int qtAlim) {
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setBold(true);
		style.setFont(font);
		for (int i = 0; i < 2 + (qtAlim * 9); i++) {
			rowHeader.getCell(i).setCellStyle(style);
		}
	}
	
	private void createStaticHeaderFields(Row rowHeader) {
		rowHeader.createCell(0).setCellValue("DATA");
		rowHeader.createCell(1).setCellValue("HORA");
	}

	private Medicao setMedicaoFromFields(String[] fields) throws ParseException {
		Medicao medicao = new Medicao();
		util.replaceNullAndEmptyFields(fields);
		medicao.setData(util.string2Calendar(fields[2]+ "00"));
		medicao.setCorrFaseA(Integer.parseInt(fields[5]));
		medicao.setCorrFaseB(Integer.parseInt(fields[6]));
		medicao.setCorrFaseC(Integer.parseInt(fields[7]));
		medicao.setPotAtiva(Integer.parseInt(fields[8]));
		medicao.setPotReat(Integer.parseInt(fields[9]));
		medicao.setFatorPot(Double.parseDouble(fields[10].replace(',', '.')));
		medicao.setTensaoA(Double.parseDouble(fields[11].replace(',', '.')));
		medicao.setTensaoB(Double.parseDouble(fields[12].replace(',', '.')));
		medicao.setTensaoC(Double.parseDouble(fields[13].replace(',', '.')));
		return medicao;
	}
	
	private void writeMedicao(Row newRow, Medicao medicao, int initialCell) {
		newRow.createCell(initialCell).setCellValue(medicao.getCorrFaseA());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getCorrFaseB());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getCorrFaseC());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getPotAtiva());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getPotReat());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getFatorPot());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getTensaoA());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getTensaoB());
		initialCell++;
		newRow.createCell(initialCell).setCellValue(medicao.getTensaoC());
	}
	
	@SuppressWarnings("unused")
	private int findInitialCell(String alimentadorCode, SXSSFSheet sheet) {
		Row headerRow = sheet.getRow(0);
		for (Cell cell : headerRow) {
			if (cell.getRichStringCellValue().getString().contains(alimentadorCode)) {
				return cell.getColumnIndex();
			}
		}
		return -1;
	}

	private HashMap<Integer, String[]> createReferenceMap(String referenceFile) throws IOException {
		
		HashMap<Integer, String[]> referenceMap = new HashMap<>();
		br = util.openBufferedReader(referenceFile);
		
		br.readLine();
		
		String line;
		while ((line = br.readLine()) != null) {
			String[] referenceFields = line.split(";");
			String[] alimentadorData = {referenceFields[3], referenceFields[2]};
			referenceMap.put(Integer.parseInt(referenceFields[4]), alimentadorData);
		}
		
		return referenceMap;
	}
	
}
