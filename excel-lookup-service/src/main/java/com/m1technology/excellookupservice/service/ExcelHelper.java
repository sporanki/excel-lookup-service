package com.m1technology.excellookupservice.service;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.web.multipart.MultipartFile;

import com.google.gson.Gson;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelHelper {
	public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	public static boolean hasExcelFormat(MultipartFile file) {

		if (!TYPE.equals(file.getContentType())) {
			return false;
		}

		return true;
	}

	private static final String XLS = "xls";
	private static final String XLSX = "xlsx";

	public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
		Workbook workbook = null;
		if (fileType.equalsIgnoreCase(XLS)) {
			workbook = new HSSFWorkbook(inputStream);
		} else if (fileType.equalsIgnoreCase(XLSX)) {
			workbook = new XSSFWorkbook(inputStream);
		}
		return workbook;
	}

	/*
	 * Read data from an excel file and output each sheet data to a json file and a
	 * text file. filePath : The excel file store path.
	 */
	public static String creteJSONAndTextFileFromExcel(MultipartFile file) {
		String jsonString = null;
		/* First need to open the file. */
		// FileInputStream fInputStream = new FileInputStream(filePath.trim());

		try {
			String fileName = file.getOriginalFilename();
			
			if (file == null || file.isEmpty() || fileName.lastIndexOf(".") < 0) {
				log.error("Excel , Errors！");
				return null;
			}
			

			String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());

			FileInputStream fInputStream = (FileInputStream) file.getInputStream();

			/* Create the workbook object to access excel file. */
			// Workbook excelWookBook = new XSSFWorkbook(fInputStream)
			/*
			 * Because this example use .xls excel file format, so it should use
			 * HSSFWorkbook class. For .xlsx format excel file use XSSFWorkbook class.
			 */;
			Workbook excelWorkBook = getWorkbook(fInputStream, fileType);

			// Get all excel sheet count.
			int totalSheetNumber = excelWorkBook.getNumberOfSheets();

			// Loop in all excel sheet.
			for (int i = 0; i < totalSheetNumber; i++) {
				// Get current sheet.
				Sheet sheet = excelWorkBook.getSheetAt(i);

				// Get sheet name.
				String sheetName = sheet.getSheetName();

				if (sheetName != null && sheetName.length() > 0) {
					// Get current sheet data in a list table.
					List<List<Object>> sheetDataTable = getSheetDataList(sheet);

					// Generate JSON format of above sheet data and write to a JSON file.
					jsonString = getJSONStringFromList(sheetDataTable);
					// String jsonFileName = sheet.getSheetName() + ".json";
					// writeStringToFile(jsonString, jsonFileName);

				}
			}
			// Close excel work book object.
			excelWorkBook.close();
		} catch (Exception ex) {
			System.err.println(ex.getMessage());
		}

		return jsonString;
	}

	/*
	 * Return sheet data in a two dimensional list. Each element in the outer list
	 * is represent a row, each element in the inner list represent a column. The
	 * first row is the column name row.
	 */
	private static List<List<Object>> getSheetDataList(Sheet sheet) {
		List<List<Object>> ret = new ArrayList<List<Object>>();

		// Get the first and last sheet row number.
		int firstRowNum = sheet.getFirstRowNum();
		int lastRowNum = sheet.getLastRowNum();

		if (lastRowNum > 0) {
			
			Row header = sheet.getRow(0);
	        if (header == null) {
	            log.warn("Header is Missing");
	            return null;
	        }
	        
			// Loop in sheet rows.
			for (int i = firstRowNum; i < lastRowNum + 1; i++) {
				// Get current row object.
				Row row = sheet.getRow(i);

				// Get first and last cell number.
				int firstCellNum = row.getFirstCellNum();
				int lastCellNum = row.getLastCellNum();

				// Create a String list to save column data in a row.
				List<Object> rowDataList = new ArrayList<Object>();

				// Loop in the row cells.
				for (int j = firstCellNum; j < lastCellNum; j++) {
					// Get current cell.
					Cell cell = row.getCell(j);

					// Get cell type.
					CellType cellType = cell.getCellType();
					 Object value = null;
			        switch (cellType) {
		            case NUMERIC:
		                Double doubleValue = cell.getNumericCellValue();
		                DecimalFormat decimalFormat = new DecimalFormat("0");
		                value = decimalFormat.format(doubleValue);
		                
		                rowDataList.add(value);
		                //if (type.getTypeName().equals("long")) value = Long.valueOf((String) value);
		                break;
		            case STRING:
		                String str = cell.getStringCellValue().trim();
		                if (str.equals("")) break;
		                value = str;
		                rowDataList.add(value);
		                break;
		            case BOOLEAN:
		                value = cell.getBooleanCellValue();
		                rowDataList.add(value);
		                break;
		            case FORMULA:
		                value = cell.getCellFormula();
		                rowDataList.add(value);
		                break;
		            case BLANK:
		            	rowDataList.add("");
		            	break;
		            case _NONE:
		                break;
		            default:
		                break;
		        }
				}

				// Add current row data list in the return list.
				ret.add(rowDataList);
			}
		}
		return ret;
	}

	/* Return a JSON string from the string list. */
	private static String getJSONStringFromList(List<List<Object>> dataTable) {
		//String ret = "";
		JSONArray att = new JSONArray();
		Gson gson = new Gson();
		if (dataTable != null) {
			int rowCount = dataTable.size();
			
			if (rowCount > 1) {
				// Create a JSONObject to store table data.
				//JSONObject tableJsonObject = new JSONObject();
				
				// Gson gson = new Gson();
				// String jsonString = gson.toJson(list);

				// The first row is the header row, store each column name.
				List<Object> headerRow = dataTable.get(0);

				int columnCount = headerRow.size();

				// Loop in the row data list.
				for (int i = 1; i < rowCount; i++) {
					// Get current row data.
					List<Object> dataRow = dataTable.get(i);

					if (!dataRow.isEmpty()) {

						// Create a JSONObject object to store row data.
						//Gson rowGson = new Gson();
						JSONObject rowJsonObject = new JSONObject();

						for (int j = 0; j < columnCount; j++) {
							String columnName = (String) headerRow.get(j);
							String columnValue = (String) dataRow.get(j);
							rowJsonObject.put(columnName, columnValue);
						}
						System.out.println("Row " + i + rowJsonObject);
						att.put(rowJsonObject);
						//tableJsonObject.put("Row " + i, rowJsonObject);
					}
				}
			}
		}
		return new Gson().toJson(att.toList());
	}

}
