package atilla.tech.ex4j.service;

import java.util.HashMap;
import java.util.Map;

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
			//Assume the first row as the title row.
			XSSFRow titleRow = worksheet.getRow(0);
			
			//Get column names
			Map<String, Integer> columns = new HashMap<>();
			titleRow.forEach(cell -> {
				columns.put(cell.getStringCellValue(), cell.getColumnIndex());
			});
			
			//This json object will be returned at the end.
			JSONObject items = new JSONObject();
			//start to read from the second row since the first row assumed as title row.
			for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
				XSSFRow row = worksheet.getRow(i);
				//this is a single object json
				JSONObject item = new JSONObject();
				for(String columnName : columns.keySet()) {
					Integer columnNo = columns.get(columnName);
					XSSFCell cell = row.getCell(columnNo);
					DataFormatter formatter = new DataFormatter();
					String cellValue = formatter.formatCellValue(cell);
					
					item.put(columnName, cellValue);
				}
				items.put(i+"", item);
			}
			workbook.close();
			return items.toString();
		} catch (Exception e) {
			JSONObject json = new JSONObject();
			json.put("message", e.getMessage());
			json.put("cause", e.getCause());
			return json.toString();
		}
	}

}
