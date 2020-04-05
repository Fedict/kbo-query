/*
 * Copyright (c) 2020, Bart Hanssens <bart.hanssens@bosa.fgov.be>
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * * Redistributions of source code must retain the above copyright notice, this
 *   list of conditions and the following disclaimer.
 * * Redistributions in binary form must reproduce the above copyright notice,
 *   this list of conditions and the following disclaimer in the documentation
 *   and/or other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
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
import java.nio.file.Paths;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Custom query for mapping CBE open data
 * 
 * @author Bart Hanssens
 */
public class Main {
	/**
	 * 
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		if (args.length < 3) {
			System.err.println("usage: input-file directory-kbo-csv output-file");
			System.exit(-1);
		}

		String input = args[0];
		String kboDir = args[1];
		String output = args[2];
		
		List<Row> rows = readInput(input);
		Set<String> toCheck = getNrsToCheck(rows);

		Map<String,String> naceNL = readCodes(kboDir, "code.csv", "NL");
		Map<String,String> naceFR = readCodes(kboDir, "code.csv", "FR");

		Multimap<String,String> vatActMain = readActivities(toCheck, kboDir, "activity.csv", "BTW001", "MAIN");
		Multimap<String,String> vatActSeco = readActivities(toCheck, kboDir, "activity.csv", "BTW001", "SECO");
		Multimap<String,String> hActMain = readActivities(toCheck, kboDir, "activity.csv", "RSZ001", "MAIN");
		Multimap<String,String> hActSeco = readActivities(toCheck, kboDir, "activity.csv", "RSZ001", "SECO");

		writeOutput(kboDir, output, rows, vatActMain, vatActSeco, hActMain, hActSeco, naceNL, naceFR);
	}
	
	/**
	 * Read input file (XLSX) and keep the rows in memory
	 * 
	 * @param inFile input file
	 * @return list of rows
	 * @throws IOException 
	 */
	private static List<Row> readInput(String inFile) throws IOException {
		List<Row> rows = new ArrayList<>();
		
		try (InputStream is = Files.newInputStream(Paths.get(inFile))) {
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
	 * Write result to output file (XLSX)
	 * 
	 * @param outDir output directory
	 * @param file output file
	 * @param rows rows
	 * @param vatActMain VAT main activity
	 * @param vatActSeco VAT secondary activity
	 * @param hActMain NSSO main activity
	 * @param hActSeco NSSO secondary activity
	 * @param naceNL NACE codes Dutch
	 * @param naceFR NACE code French
	 * @throws IOException 
	 */
	private static void writeOutput(String outDir, String file, List<Row> rows, 
			Multimap<String,String> vatActMain,Multimap<String,String> vatActSeco,
			Multimap<String,String> hActMain,Multimap<String,String> hActSeco,
			Map<String,String> naceNL, Map<String,String> naceFR) throws IOException {
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet();
		CellStyle cs = wb.createCellStyle();
		cs.setWrapText(true);

		for (int i = 0; i < rows.size(); i++) {
			Row newRow = sheet.createRow(i);

			// copy existing rows			
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
			
			// add new columns about activities
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

		try (OutputStream os = Files.newOutputStream(Paths.get(outDir, file))) {
			wb.write(os);
		}
		wb.close();
	}
	
	/**
	 * Get company numbers from excel
	 * 
	 * @param rows
	 * @return set of company numbers
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

	/**
	 * Get the VAT/NSSO activities for a list of companies
	 * 
	 * @param nrs companies to check
	 * @param dir input directory
	 * @param file activity file
	 * @param actGroup activity group (VAT or NSSO)
	 * @param importance main or secondary activity
	 * @return
	 * @throws IOException 
	 */
	private static Multimap<String,String> readActivities(Set<String> nrs, String dir, String file, 
																String actGroup, String importance) throws IOException {
		SetMultimap<String,String> multi = HashMultimap.create();
		try (	Reader reader = Files.newBufferedReader(Paths.get(dir, file));
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

	/**
	 * Read NACE code table and create a map of NACE numbers and label in specific language
	 * 
	 * @param dir directory
	 * @param file file to read
	 * @param lang language
	 * @return map of code and label
	 * @throws IOException 
	 */
	private static Map<String, String> readCodes(String dir, String file, String lang) throws IOException {
		HashMap<String,String> codes = new HashMap<>(2048);
		try (	Reader reader = Files.newBufferedReader(Paths.get(dir, file));
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

	/**
	 * 
	 * @param nrs
	 * @param dir
	 * @param file
	 * @return
	 * @throws IOException 
	 */
	private static Multimap<String,String> readSites(Set<String> nrs, String dir, String file) throws IOException {
		SetMultimap<String,String> multi = HashMultimap.create();
		try (	Reader reader = Files.newBufferedReader(Paths.get(dir, file));
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
