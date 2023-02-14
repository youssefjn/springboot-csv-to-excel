package excel.csv.excelcsv.controller;

import java.io.IOException;


import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import excel.csv.excelcsv.service.ConverterService;

@RestController
public class ConverterController {
    @Autowired
  private ConverterService converterService;
@PostMapping(value="/CsvToExcel")
    public ResponseEntity<String> CsvToExcel(@RequestPart MultipartFile csvFile) throws IOException{
        converterService.convertCsvToExcel(csvFile);
return new ResponseEntity<>("done", HttpStatus.OK);
    }
}
