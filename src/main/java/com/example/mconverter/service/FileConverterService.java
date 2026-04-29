package com.example.mconverter.service;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

import java.io.ByteArrayOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

@Service
public class FileConverterService {

public byte[] wordToPdf(MultipartFile file) throws Exception {
    // Alternative avec iText (extraction du texte puis création PDF)
    XWPFDocument document = new XWPFDocument(file.getInputStream());
    
    // Extraire tout le texte du document Word
    StringBuilder text = new StringBuilder();
    for (XWPFParagraph paragraph : document.getParagraphs()) {
        text.append(paragraph.getText()).append("\n");
    }
    
    // Extraire aussi le texte des tableaux si présents
    for (XWPFTable table : document.getTables()) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                text.append(cell.getText()).append("\t");
            }
            text.append("\n");
        }
    }
    document.close();
    
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    com.itextpdf.text.Document pdfDoc = new com.itextpdf.text.Document();
    com.itextpdf.text.pdf.PdfWriter.getInstance(pdfDoc, out);
    pdfDoc.open();
    
    com.itextpdf.text.Paragraph paragraph = new com.itextpdf.text.Paragraph(text.toString());
    pdfDoc.add(paragraph);
    
    pdfDoc.close();
    
    return out.toByteArray();
}

    public byte[] pdfToWord(MultipartFile file) throws Exception {
        // Implémentation PDF -> Word
        PDDocument pdfDoc = PDDocument.load(file.getInputStream());
        PDFTextStripper stripper = new PDFTextStripper();
        String text = stripper.getText(pdfDoc);
        pdfDoc.close();

        // Créer un document Word avec le texte
        XWPFDocument wordDoc = new XWPFDocument();
        wordDoc.createParagraph().createRun().setText(text);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wordDoc.write(out);
        wordDoc.close();

        return out.toByteArray();
    }

    public byte[] pdfToExcel(MultipartFile file) throws Exception {
        // Implémentation PDF -> Excel
        PDDocument pdfDoc = PDDocument.load(file.getInputStream());
        PDFTextStripper stripper = new PDFTextStripper();
        String text = stripper.getText(pdfDoc);
        pdfDoc.close();

        // Créer un fichier Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("PDF Content");

        String[] lines = text.split("\n");
        for (int i = 0; i < lines.length; i++) {
            XSSFRow row = sheet.createRow(i);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(lines[i]);
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();

        return out.toByteArray();
    }
}