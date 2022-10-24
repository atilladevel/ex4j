package atilla.tech.ex4j.web.rest;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import atilla.tech.ex4j.service.ExcelService;

@RestController
@RequestMapping("/excel/")
public class ExcelController {

	@Autowired
	private ExcelService excelService;
	
	@RequestMapping(value="read/sheet/{sheetNo}", method = RequestMethod.POST)
	public String readExcelSheet(@PathVariable Integer sheetNo, @RequestParam("excelFile") MultipartFile file) {
		return excelService.readExcelSheet(file, sheetNo);
	}
	
	@RequestMapping(value="write/list", method = RequestMethod.POST)
	public void writeExcelFile(@RequestBody String data) {
		excelService.writeExcelSheet(data);
	}

}
