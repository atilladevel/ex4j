package atilla.tech.ex4j.service;

import java.io.FileOutputStream;
import java.time.Instant;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelService {

	public String readExcelSheet(MultipartFile file, Integer sheetNo) {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
			XSSFSheet worksheet = workbook.getSheetAt(sheetNo);
			// Assume the first row as the title row.
			XSSFRow titleRow = worksheet.getRow(0);

			// Get column names
			Map<String, Integer> columns = new HashMap<>();
			titleRow.forEach(cell -> {
				columns.put(cell.getStringCellValue(), cell.getColumnIndex());
			});

			// This json object will be returned at the end.
			JSONObject items = extractSheetData(worksheet, columns);
			workbook.close();
			return items.toString();
		} catch (Exception e) {
			JSONObject json = new JSONObject();
			json.put("error", e.getMessage());
			return json.toString();
		}
	}

	public void writeExcelSheet(String jsonData) {
		try {
			JSONObject data = new JSONObject(jsonData);
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet();
			
			int rowCounter = 0;
			Iterator<String> itemIterator = data.keys();
			while(itemIterator.hasNext()) {
				String itemIndex = itemIterator.next();
				if(data.get(itemIndex) instanceof JSONObject) {
					int cellCounter = 0;
					XSSFRow row = sheet.createRow(rowCounter);
					JSONObject item = data.getJSONObject(itemIndex);
					Iterator<String> columnIterator = item.keys();
					while(columnIterator.hasNext()) {
						String columnName = columnIterator.next();
						XSSFCell cell = row.createCell(cellCounter, CellType.STRING);
						cell.setCellValue(item.getString(columnName));
						cellCounter++;
					}
				}
				rowCounter++;
			}
			
			String filename = Instant.now().toString() + ".xlsx";
			FileOutputStream outputStream = new FileOutputStream(filename);
	        workbook.write(outputStream);
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private JSONObject extractSheetData(XSSFSheet sheet, Map<String, Integer> columns) {
		JSONObject items = new JSONObject();
		// start to read from the second row since the first row assumed as title row.
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow row = sheet.getRow(i);
			// this is a single object json
			JSONObject item = new JSONObject();
			for (String columnName : columns.keySet()) {
				Integer columnNo = columns.get(columnName);
				XSSFCell cell = row.getCell(columnNo);
				DataFormatter formatter = new DataFormatter();
				String cellValue = formatter.formatCellValue(cell);

				item.put(columnName, cellValue);
			}
			items.put(i + "", item);
		}
		return items;
	}
}
