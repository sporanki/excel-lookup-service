package com.m1technology.excellookupservice.controllers;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.m1technology.excellookupservice.model.ResponseMessage;
import com.m1technology.excellookupservice.service.ExcelHelper;


@CrossOrigin("http://localhost:8080")
@Controller
@RequestMapping("/api/excel")
public class ExcelController {


  @PostMapping("/upload")
  public ResponseEntity<?> uploadFile(@RequestParam("file") MultipartFile file) {
    String message = "";
    
    try {
    	message = ExcelHelper.creteJSONAndTextFileFromExcel(file);

        //message = "Uploaded the file successfully: " + file.getOriginalFilename();
        //return ResponseEntity.status(HttpStatus.OK).body(new ResponseMessage(message));
    	return ResponseEntity.status(HttpStatus.OK).body(message);
      } catch (Exception e) {
        message = "Could not upload the file: " + file.getOriginalFilename() + "!";
        return ResponseEntity.status(HttpStatus.EXPECTATION_FAILED).body(new ResponseMessage(message));
      }
    

   // message = "Please upload an excel file!";
    //return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(new ResponseMessage(message));
  }

  @GetMapping("/hello")
  public ResponseEntity<ResponseMessage> hello() {
    String message = "";
    
        message = "Uploaded the file successfully: Hello ";
        return ResponseEntity.status(HttpStatus.OK).body(new ResponseMessage(message));
  }
  
}