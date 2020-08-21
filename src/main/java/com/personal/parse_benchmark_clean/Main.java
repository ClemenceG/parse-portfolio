package com.personal.parse_benchmark_clean;



	import java.io.File;

	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.io.FileNotFoundException;
	import java.io.IOException;
	import java.util.Iterator;
import java.util.Scanner;
import java.util.ArrayList;

	import org.apache.poi.ss.usermodel.DateUtil;
	import java.util.Date;

	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	
	public class Main {
		
		private static String benchmarkName = "Eurostoxx";
		private static String[] columns = {"ISIN", "Weight", "Currency", "SuperSector Nb ", "SuperSector Name", "Sector Nb", "Sector Name", "Subsector Nb", "Subsector Name", "Country Code", "Country Name"};
		
		public static void main(String[] args) {
			Scanner scanner = new Scanner(System.in);
			System.out.println("Prendre comme example l'onglet \"copie en dur\" du fichier Teams: Product/05-Projects/03-Performance/Portfolio Data/Parsing Template/Eurostoxx.xlsx:");
			System.out.println("Entrez path/filename du fichier à extraire :");
			String filename = scanner.nextLine();
			
			System.out.println("Entrez path/filename du fichier de destination :");
			String exportFilename = scanner.nextLine();
			
			//System.out.println("Souhitez-vous extraire les transactions? oui/non");
			//boolean parse_transactions = scanner.nextLine().contentEquals("oui");
			scanner.close();
			
			
			File xlsx = new File(filename);
			try(FileInputStream is = new FileInputStream(xlsx);
					XSSFWorkbook wb = new XSSFWorkbook(is);) {
				
				XSSFSheet sheet = wb.getSheetAt(0);
				Benchmark benchmark = parseBenchmarkComposition(sheet);
				
				sheet = wb.getSheetAt(1);
				parsePriceHistory(benchmark, sheet);
				
				sheet = wb.getSheetAt(2);
				parseCharacteristicSheet(sheet, benchmark);
				
				printExcelFile(exportFilename, benchmark);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
		/**
		 * Method returns the benchmark obtained from the first sheet
		 * calls 1 submethods
		 * @param wb
		 * @return
		 */
		private static Benchmark parseBenchmarkComposition(XSSFSheet sheet) {

			Benchmark benchmark = new Benchmark(benchmarkName);
			
			Iterator<Row> rowIterator = sheet.iterator();
			
			// skip header line
			Row currentRow = rowIterator.next();
			
			while(rowIterator.hasNext()) {
				currentRow = rowIterator.next();
				BenchmarkElement benchmarkElement = parseBenchmarkElementFromRow(currentRow);
				benchmark.addBenchmarkElement(benchmarkElement);
			}
			return benchmark;
		}
		
		/**
		 * obtains the benchmark element from a line
		 * Sub-method for parsing the first sheet
		 * @param currentRow
		 * @return
		 */
		private static BenchmarkElement parseBenchmarkElementFromRow(Row currentRow) {
			Iterator<Cell> cellIterator = currentRow.iterator();
			
			int columnIndex = 1;
			
			// initialize elements of the benchmark element to read
			String name = null;
			double weight = 0;
			String ticker = null;
			
			while(cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				switch(columnIndex) {
				case 1:
					name = currentCell.getStringCellValue();
					break;
				case 2:
					weight = currentCell.getNumericCellValue();
					break;
				case 3:
					ticker = currentCell.getStringCellValue();
					break;
				}
				columnIndex++;
			}
			Security security = new Security(name, ticker);
			return new BenchmarkElement(security, weight);
		}
		
		/**
		 * Method that exports and associates the price history from the second sheet to the securities
		 * @param benchmark
		 * @param sheet
		 * @param nameDivider
		 */
		private static void parsePriceHistory(Benchmark benchmark, XSSFSheet sheet) {
			
			Iterator<Row> rowIterator = sheet.iterator();
			
			// skip header line
			Row currentRow = rowIterator.next();
			int benchmarkIndex = 0;
			
			while(rowIterator.hasNext()) {
				
				currentRow = rowIterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				
				while(cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();

					if(currentCell.getStringCellValue().equals(benchmark.getSecurityBBGlobal(benchmarkIndex))) {
						currentCell = cellIterator.next();
						Date date = DateUtil.getJavaDate(currentCell.getNumericCellValue());
						currentCell = cellIterator.next();
						double price = currentCell.getNumericCellValue();
						PriceElement priceElement = new PriceElement(date, price);
						benchmark.getSecurity(benchmarkIndex).addPriceElement(priceElement);
					} else {
						benchmarkIndex++;
						// skip the start line
						if(rowIterator.hasNext()) {
							currentRow = rowIterator.next();
						}
						break;
					}
				}
				
			}
		}
		
		private static void parseCharacteristicSheet(XSSFSheet sheet, Benchmark benchmark) {
			Iterator<Row> rowIterator = sheet.iterator();
			
			// skip header line
			Row currentRow = rowIterator.next();

			
			while(rowIterator.hasNext()) {
				int rowIndex = currentRow.getRowNum();
				Security currentSecurity = benchmark.getSecurity(rowIndex);
				
				currentRow = rowIterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				
				while(cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					int cellIndex = currentCell.getColumnIndex();
					
					switch(cellIndex) {
					case 0:
						break;
					case 1:
						break;
					case 2:
						break;
					case 3:
						currentSecurity.setCurrency(currentCell.getStringCellValue());
						break;
					case 4:
						Sector sector = new Sector(currentCell.getStringCellValue());
						currentSecurity.setSector(sector);
						break;
					case 5:
						currentSecurity.getSector().addSectorNb((int) currentCell.getNumericCellValue());
						break;
					case 6:
						sector = new Sector(currentCell.getStringCellValue());
						currentSecurity.setSubsector(sector);
						break;
					case 7:
						currentSecurity.getSubsector().addSectorNb((int) currentCell.getNumericCellValue());
						break;
					case 8:
						sector = new Sector(currentCell.getStringCellValue());
						currentSecurity.setSupersector(sector);
						break;
					case 9:
						currentSecurity.getSupersector().addSectorNb((int) currentCell.getNumericCellValue());
						break;

						// 10-15 GICS Info missing
					case 18:
						Country country = new Country(currentCell.getStringCellValue());
						currentSecurity.setCountry(country);
						break;
					case 19:
						currentSecurity.getCountry().setCountryIso(currentCell.getStringCellValue());
						break;
					case 22:
						currentSecurity.setIsin(currentCell.getStringCellValue());
						break;
					case 23:
						currentSecurity.setTicker(currentCell.getStringCellValue());
						break;
						
					}
					
				}
			}
		}
		
		private static void printExcelFile(String filename, Benchmark benchmark) {
			
			try(FileOutputStream fileOut = new FileOutputStream(new File(filename));
					Workbook wb = new XSSFWorkbook();) {

				Sheet sheet = wb.createSheet("Eurostoxx");
				CreationHelper creationHelper = wb.getCreationHelper();

				// initialize header row
				Row headerRow = sheet.createRow(0);
				int i = 0;
				for(; i < columns.length; i++) {
					Cell cell = headerRow.createCell(i);
					cell.setCellValue(columns[i]);
				}

				ArrayList<Date> dateHeader = parseDateHeader(benchmark);
				
				
				for(Date date : dateHeader) {
					Cell cell = headerRow.createCell(i);
					cell.setCellValue(date);
					CellStyle style = wb.createCellStyle();
					style.setDataFormat(creationHelper.createDataFormat().getFormat(
							"dd/mm/yyyy"));
					cell.setCellStyle(style);
					i++;
				}
				
				i = 1;		// row index
				for(BenchmarkElement currentElement : benchmark.getBenchmarkElements()) {
					Row currentRow = sheet.createRow(i);
					Security currentSec = currentElement.getSecurity();
					
					currentRow.createCell(0).setCellValue(currentSec.getIsin());
					currentRow.createCell(1).setCellValue(currentElement.getWeight());
					currentRow.createCell(2).setCellValue(currentSec.getTicker());
					currentRow.createCell(3).setCellValue(currentSec.getCurrency());
					if(currentSec.getCountry() != null) {
						currentRow.createCell(4).setCellValue(currentSec.getCountry().getCountryIso());
						currentRow.createCell(5).setCellValue(currentSec.getSupersector().getSectorNb());
						currentRow.createCell(6).setCellValue(currentSec.getSupersector().getSectorName());
						currentRow.createCell(7).setCellValue(currentSec.getSector().getSectorNb());
						currentRow.createCell(8).setCellValue(currentSec.getSector().getSectorName());
						currentRow.createCell(9).setCellValue(currentSec.getSubsector().getSectorNb());
						currentRow.createCell(10).setCellValue(currentSec.getSubsector().getSectorName());

						int sheetIndex = 11;
						int securityIndex = 0;
						for(int j=0; securityIndex<currentSec.getPrices().size(); j++) {
							PriceElement element = currentSec.getPrices().get(securityIndex);
							if(element.getDate().equals(dateHeader.get(j))) {
								currentRow.createCell(sheetIndex).setCellValue(element.getPrice());
							} else {
								currentRow.createCell(sheetIndex).setCellValue(0);
								securityIndex--;
							}
							securityIndex++;
							sheetIndex++;
						}
					}
					i++;
				}
				
				
		        wb.write(fileOut);

			} catch (Exception e) {
				e.printStackTrace();
			}
			
		}

		private static ArrayList<Date> parseDateHeader(Benchmark benchmark) {

			ArrayList<Date> dateHeader = new ArrayList<Date>();
			
			for(BenchmarkElement currentElement : benchmark.getBenchmarkElements()) {
				int i = 0;
				if(dateHeader.isEmpty()) {
					for(PriceElement currentPriceEl : currentElement.getSecurity().getPrices()) {
						Date date = currentPriceEl.getDate();
						dateHeader.add(date);
					}

				} else {
					for(PriceElement currentPriceEl : currentElement.getSecurity().getPrices()) {
						Date date = currentPriceEl.getDate();
						if(!(dateHeader.get(i).equals(date))) {
							dateHeader.add(i, date);
						}
						i++;
					}
				}
			}
			return dateHeader;
		}
		
		/*private static String debugger(Cell cell) {
			String cellValue = null;
			if(cell.getCellType().equals(CellType.NUMERIC)) {
				double d = cell.getNumericCellValue();
				cellValue = Double.toString(d);
			} else if(cell.getCellType().equals(CellType.STRING)) {
				cellValue = cell.getStringCellValue();
			}
			return cellValue;
		}*/
	}


