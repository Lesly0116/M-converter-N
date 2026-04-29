package com.example.mconverter.controller;

import com.example.mconverter.service.FileConverterService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/convert")
@CrossOrigin("*")
public class FileConverterController {

    private final FileConverterService service;

    public FileConverterController(FileConverterService service) {
        this.service = service;
    }

    @PostMapping("/word-to-pdf")
    public ResponseEntity<byte[]> wordToPdf(@RequestParam("file") MultipartFile file) {
        try {
            System.out.println("=== Word to PDF conversion started ===");
            System.out.println("File name: " + file.getOriginalFilename());
            System.out.println("File size: " + file.getSize());
            
            byte[] result = service.wordToPdf(file);
            
            System.out.println("Conversion successful, result size: " + result.length);
            
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted.pdf")
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(result);
        } catch (Exception e) {
            System.err.println("ERROR in wordToPdf: " + e.getMessage());
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }
    }

    @PostMapping("/pdf-to-word")
    public ResponseEntity<byte[]> pdfToWord(@RequestParam("file") MultipartFile file) {
        try {
            System.out.println("=== PDF to Word conversion started ===");
            System.out.println("File name: " + file.getOriginalFilename());
            System.out.println("File size: " + file.getSize());
            
            byte[] result = service.pdfToWord(file);
            
            System.out.println("Conversion successful, result size: " + result.length);
            
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted.docx")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(result);
        } catch (Exception e) {
            System.err.println("ERROR in pdfToWord: " + e.getMessage());
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }
    }

    @PostMapping("/pdf-to-excel")
    public ResponseEntity<byte[]> pdfToExcel(@RequestParam("file") MultipartFile file) {
        try {
            System.out.println("=== PDF to Excel conversion started ===");
            System.out.println("File name: " + file.getOriginalFilename());
            System.out.println("File size: " + file.getSize());
            
            byte[] result = service.pdfToExcel(file);
            
            System.out.println("Conversion successful, result size: " + result.length);
            
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted.xlsx")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(result);
        } catch (Exception e) {
            System.err.println("ERROR in pdfToExcel: " + e.getMessage());
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }
    }
}