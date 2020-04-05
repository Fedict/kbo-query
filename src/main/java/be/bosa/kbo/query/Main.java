/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package be.bosa.kbo.query;

import com.google.common.collect.HashMultimap;
import com.google.common.collect.Multimap;
import com.google.common.collect.SetMultimap;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Bart.Hanssens
 */
public class Main {
	private static Path input;
	private static Path kboDir;

	/**
	 * 
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		input = Paths.get(args[0]);
		kboDir = Paths.get(args[1]);
		
		List<Row> rows = readInput(input);
		Set<String> toCheck = getNrsToCheck(rows);

		/*
		Multimap<String,String> sites = readSites(toCheck, kboDir, "establishment.csv");
		for(Entry<String,String> entry : sites.entries()) {
			System.err.println(entry.getKey() + " " + entry.getValue());
		}
		*/

		Map<String,String> naceNL = readCodes(kboDir, "code.csv", "NL");
		Map<String,String> naceFR = readCodes(kboDir, "code.csv", "FR");

		Multimap<String,String> vatActMain = readActivities(toCheck, kboDir, "activity.csv", "BTW001", "MAIN");
		Multimap<String,String> vatActSeco = readActivities(toCheck, kboDir, "activity.csv", "BTW001", "SECO");
		Multimap<String,String> hActMain = readActivities(toCheck, kboDir, "activity.csv", "RSZ001", "MAIN");
		Multimap<String,String> hActSeco = readActivities(toCheck, kboDir, "activity.csv", "RSZ001", "SECO");

		writeOutput(kboDir, "output.xlsx", rows, vatActMain, vatActSeco, hActMain, hActSeco, naceNL, naceFR);
	}
	
	/**
	 * 
	 * @param inFile
	 * @return
	 * @throws IOException 
	 */
	private static List<Row> readInput(Path inFile) throws IOException {
		List<Row> rows = new ArrayList<>();
		
		try (InputStream is = Files.newInputStream(inFile)) {
			XSSFWorkbook myWorkBook = new XSSFWorkbook(is); 
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.iterator();

			while (rowIterator.hasNext()) { 
				rows.add(rowIterator.next());
			}
		}
		return rows;
	}
	/**
	 * 
	 * @param inFile
	 * @return
	 * @throws IOException 
	 */
	private static void writeOutput(Path outFile, String file, List<Row> rows, 
			Multimap<String,String> vatActMain,Multimap<String,String> vatActSeco,
			Multimap<String,String> hActMain,Multimap<String,String> hActSeco,
			Map<String,String> naceNL, Map<String,String> naceFR) throws IOException {
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet();
		CellStyle cs = wb.createCellStyle();
		cs.setWrapText(true);

		for (int i = 0; i < rows.size(); i++) {
			Row newRow = sheet.createRow(i);
			
			for (int j = 0; j < 9; j++) {
				Cell cell = rows.get(i).getCell(j);
				Cell newCell = newRow.createCell(j);
				if (cell != null) {
					if (cell.getCellTypeEnum().equals(CellType.NUMERIC)) {
						newCell.setCellValue(cell.getNumericCellValue());
					}
					if (cell.getCellTypeEnum().equals(CellType.STRING)) {
						newCell.setCellValue(cell.getStringCellValue());
					}
				}
			}
			Cell compcell = rows.get(i).getCell(4);
			if (compcell != null) {
				String company = compcell.getStringCellValue();
				if (company.startsWith("0.")) {
					company = company.replaceFirst("0\\.", "0");
				}

				if (vatActMain.containsKey(company)) {
					Cell newCell = newRow.createCell(10);
					newCell.setCellStyle(cs);
					Collection<String> vals = vatActMain.get(company);
					StringBuilder buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceFR.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
					
					newCell = newRow.createCell(11);
					newCell.setCellStyle(cs);
					vals = vatActMain.get(company);
					buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceNL.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
				}
				if (vatActSeco.containsKey(company)) {
					Cell newCell = newRow.createCell(12);
					newCell.setCellStyle(cs);
					Collection<String> vals = vatActSeco.get(company);
					StringBuilder buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceFR.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
					
					newCell = newRow.createCell(13);
					newCell.setCellStyle(cs);
					vals = vatActSeco.get(company);
					buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceNL.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
					
				}
				if (hActMain.containsKey(company)) {
					Cell newCell = newRow.createCell(14);
					newCell.setCellStyle(cs);
					Collection<String> vals = hActMain.get(company);
					StringBuilder buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceFR.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
					
					newCell = newRow.createCell(15);
					newCell.setCellStyle(cs);
					vals = hActMain.get(company);
					buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceNL.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
				}
				if (hActSeco.containsKey(company)) {
					Cell newCell = newRow.createCell(16);
					newCell.setCellStyle(cs);
					Collection<String> vals = hActSeco.get(company);
					StringBuilder buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceFR.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
					
					newCell = newRow.createCell(17);
					newCell.setCellStyle(cs);
					vals = hActSeco.get(company);
					buf = new StringBuilder();
					for (String val: vals) {
						buf.append(val).append(" - ").append(naceNL.get(val)).append("\n");
					}
					newCell.setCellValue(buf.toString());
				}
			}
		}

		try (OutputStream os = Files.newOutputStream(Paths.get(outFile.toString(), file))) {
			wb.write(os);
		}
		wb.close();
	}
	
	/**
	 * 
	 * @param rows
	 * @return 
	 */
	private static Set<String> getNrsToCheck(List<Row> rows) {
		Set<String> nrs = new HashSet<>(2048);
		for (Row row: rows) {
			Cell cell = row.getCell(4);
			if (cell != null) {
				String company = cell.getStringCellValue();
				if (company.startsWith("0.")) {
					company = company.replaceFirst("0\\.", "0");
				}
				nrs.add(company);
			}
		}
		
		return nrs;
	}
	
	private static Multimap<String,String> readActivities(Set<String> nrs, Path dir, String file, 
																String actGroup, String importance) throws IOException {
		SetMultimap<String,String> multi = HashMultimap.create();
		try (	Reader reader = Files.newBufferedReader(Paths.get(dir.toString(), file));
				CSVReader csvReader = new CSVReader(reader)) {
			String[] row;
            while ((row = csvReader.readNext()) != null) {
				if (row[1].equals(actGroup) && row[2].equals("2008") && row[4].equals(importance) && nrs.contains(row[0])) {
					multi.put(row[0], row[3]);
				}
			}
		} catch (CsvValidationException ex) {
			throw new IOException(ex);
		}
		return multi;
	}

	private static Map<String, String> readCodes(Path dir, String file, String lang) throws IOException {
		HashMap<String,String> codes = new HashMap<>(2048);
		try (	Reader reader = Files.newBufferedReader(Paths.get(dir.toString(), file));
				CSVReader csvReader = new CSVReader(reader)) {
			String[] row;
            while ((row = csvReader.readNext()) != null) {
				if (row[0].equals("Nace2008") && row[2].equals(lang)) {
					codes.put(row[1], row[3]);
				}
			}
		} catch (CsvValidationException ex) {
			throw new IOException(ex);
		}
		return codes;
	}

	private static Multimap<String,String> readSites(Set<String> nrs, Path dir, String file) throws IOException {
		SetMultimap<String,String> multi = HashMultimap.create();
		try (	Reader reader = Files.newBufferedReader(Paths.get(dir.toString(), file));
				CSVReader csvReader = new CSVReader(reader)) {
			String[] row;
            while ((row = csvReader.readNext()) != null) {
				if (nrs.contains(row[2])) {
					multi.put(row[2], row[0]);
				}
			}
		} catch (CsvValidationException ex) {
			throw new IOException(ex);
		}
		return multi;
	}
}
