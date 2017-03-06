package br.com.convert;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OriginalConvertXSL {

	static DataFormatter formatter = new DataFormatter();
	static Map<String, Object> hashMap = new HashMap<>();

	static final String CELL_A = "A";
	static final String CELL_B = "B";
	static final String CELL_D = "D";
	static final String CELL_F = "F";
	static final String CELL_G = "G";
	static Map<String, String> hashData;
	static BufferedWriter writer;
	static FileWriter fw = null;
	static File files = new File("C:/Users/Lucas/Documents/Publicacoes Solr/teste.txt");

	public static void main(String args[]) {

		Map<?, ?> mapAux = convertXSL();
		createUpdate(mapAux);

	}

	public static Map<?, ?> convertXSL() {

		FileInputStream execelFile = null;

		Workbook cs = null;
		String[] textWithSplit = null;
		String id = "";

		try {
			execelFile = new FileInputStream("C:/Users/Lucas/Documents/APOIO2.xlsx");
			cs = new XSSFWorkbook(execelFile);
			Sheet dataTpe = cs.getSheetAt(0);
			Iterator<Row> rowsIteratior = dataTpe.iterator();

			while (rowsIteratior.hasNext()) {

				Row currentRow = rowsIteratior.next();

				Iterator<Cell> cellIterator = currentRow.cellIterator();
				hashData = new HashMap<>();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					if (cell.getAddress().formatAsString().startsWith(CELL_A)) {
						String text = formatter.formatCellValue(cell);
						id = text;
					}

					if (cell.getAddress().formatAsString().startsWith(CELL_B)) {
						String text = formatter.formatCellValue(cell);
						textWithSplit = text.split(",");
						if (textWithSplit.length == 4) {
							System.out.println("Tamanho do Split " +textWithSplit.length+ "  ID_SHOP " + id);
							
							hashData.put("ENDERECO", textWithSplit[0] + " " + textWithSplit[1]+" "+textWithSplit[2]);
							hashData.put("CITY", textWithSplit[2]);

							continue;
						}
						
						if(textWithSplit.length == 5){
							hashData.put("ENDERECO", textWithSplit[0] + " " + textWithSplit[1]);
							hashData.put("CITY", textWithSplit[3]);
						}
						
						if(textWithSplit.length == 6){
							hashData.put("ENDERECO", textWithSplit[0] + " " + textWithSplit[1]+" "+textWithSplit[2]);
							hashData.put("CITY", textWithSplit[3]);
						}
						if(textWithSplit.length == 1){
							hashData.put("ENDERECO", textWithSplit[0]);
							hashData.put("CITY", textWithSplit[0]);
						}
						
						if(textWithSplit.length == 7){
							hashData.put("ENDERECO", textWithSplit[0]);
							hashData.put("CITY", textWithSplit[1]);
						}
						
					}

					if (cell.getAddress().formatAsString().startsWith(CELL_D)) {
						hashData.put("PHONE", formatter.formatCellValue(cell));
					}

					if (cell.getAddress().formatAsString().startsWith(CELL_F)) {
						hashData.put("LATITUDE", formatter.formatCellValue(cell));
					}

					if (cell.getAddress().formatAsString().startsWith(CELL_G)) {
						hashData.put("LONGITUDE", formatter.formatCellValue(cell));
					}

					if (hashData.get("ENDERECO") != null && hashData.get("PHONE") != null
							&& hashData.get("LATITUDE") != null && hashData.get("LONGITUDE") != null
							&& hashData.get("CITY") != null) {

						hashMap.put(id, hashData);

					}

				}

			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println(e.getMessage());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println(e.getMessage());

			e.printStackTrace();
		}

		return hashMap;

	}

	public static void createUpdate(Map<?, ?> hash) {

		try {
			writer = new BufferedWriter(new FileWriter(files));
			for (Object key : hash.keySet()) {

				HashMap<String, String> mapQuery = (HashMap<String, String>) hash.get(key);

				String query = "UPDATE ECSL_SHOP SET ADDRESS ='" + mapQuery.get("ENDERECO") + "' , " + "PHONE_NR ='"
						+ mapQuery.get("PHONE") + "'," + "STORE_ATTRIBUTES_XML = '"
						+ "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + "<store_attributes><store_attribute id=\"1\">"
						+ "<value1>0</value1>" + "<value2>0</value2>" + "<value3>0</value3>" + "<value4>0</value4>"
						+ "<value5>" + mapQuery.get("LATITUDE") + "</value5>" + "<value6>" + mapQuery.get("LONGITUDE")
						+ "</value6>" + "<value7>0</value7> " + "</store_attribute>"
						+ "</store_attributes>'  WHERE STORE_ID = 'SGHMX' AND SHOP_ID= " + key;

				writer.newLine();
				writer.append(query);
				writer.flush();
			}
		} catch (IOException e) {
			System.out.println(e.getMessage());// TODO Auto-generated catch block
			e.printStackTrace();
			
		}
	}
}
