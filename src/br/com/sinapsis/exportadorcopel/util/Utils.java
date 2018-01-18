package br.com.sinapsis.exportadorcopel.util;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Locale;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class Utils {

	private static SimpleDateFormat sdf;
	private static SimpleDateFormat sdf1;
	private static SimpleDateFormat sdf2;
	private Calendar cal;
	
	public Utils() {
		sdf = new SimpleDateFormat("HH:mm");
		sdf1 = new SimpleDateFormat("yyyyMMddHHmm", Locale.US);
		sdf2 = new SimpleDateFormat("dd/MM/yyyy");
	}

	public Calendar string2Calendar(String field) throws ParseException {
		cal = Calendar.getInstance();
		try {
			cal.setTime(sdf1.parse(field + "00"));
		} catch (ParseException e) {
			e.printStackTrace();
			throw new ParseException("Error while parsing the date.", 0);
		}
		return cal;
	}

	public String formatDatePattern(Calendar date) {
		return sdf2.format(date.getTime());
	}

	public String formatHourPattern(Calendar date) {
		return sdf.format(date.getTime());
	}

	/**
	 * Método que varre as linhas de uma planilha e verifica se ja existe alguma
	 * linha com a data lida. Se existir retorna o número da linha da planilha,
	 * se não existe retorna -1.
	 * 
	 * @param sheet
	 *            (XSSFSheet), planilha a ser varrida
	 * @param dataMedicao
	 *            (data a ser procurada)
	 * @return int, número da linha em que existe a data procurada OU -1 se não
	 *         existir.
	 */
	public int findRow(SXSSFSheet sheet, String dateToFind, String hourToFind) {
		for (Row row : sheet) {
			if (row.getCell(0).getStringCellValue().equals(dateToFind)) {
				if (row.getCell(1).getStringCellValue().equals(hourToFind)) {
					return row.getRowNum();
				}
			}
		}
		return -1;
	}

	public void hasHeader(boolean hasHeader, BufferedReader br) throws IOException {
		if (hasHeader) {
			try {
				br.readLine();
			} catch (IOException e) {
				e.printStackTrace();
				throw new IOException("ERROR - No lines found.");
			}
		}
	}

	public void save(SXSSFWorkbook workbook, String excelPath) throws IOException {
		try {
			FileOutputStream outputStream = new FileOutputStream(new File(excelPath));
			workbook.write(outputStream);
			outputStream.close();
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException("Destiny path not found.");
		} catch (IOException e) {
			throw new IOException("Error when editing the file.");
		}
	}

	public void replaceNullAndEmptyFields(String[] fields) {
		for (int i = 0; i < fields.length; i++) {
			if (fields[i] == null || fields[i].equals("")) {
				fields[i] = "0";
			}
		}
	}

	public PrintWriter createPrintWriter(String orderedDsvPath) throws IOException {
		FileWriter outputStream;
		try {
			outputStream = new FileWriter(orderedDsvPath);
		} catch (IOException e) {
			e.printStackTrace();
			throw new IOException("ERROR - Unable to create or open the new DSV file in this path: " + orderedDsvPath);
		}
		return new PrintWriter(outputStream);
	}

	public BufferedReader openBufferedReader(String filePath) throws FileNotFoundException {
		FileReader inputStream = null;
		try {
			inputStream = new FileReader(filePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new FileNotFoundException("DSV file not found.");
		}
		return new BufferedReader(inputStream);
	}

}