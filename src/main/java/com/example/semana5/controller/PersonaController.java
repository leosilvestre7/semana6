/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.example.semana5.controller;

import com.example.semana5.service.PersonaService;
import com.example.semana5.model.Persona;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import jakarta.servlet.http.HttpServletResponse;

// importaciones pdf
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;

// importaciones excel
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.List;

@Controller
@RequestMapping("/personas")
public class PersonaController {
    
    private final PersonaService service;
    
    public PersonaController(PersonaService personaService) {
        this.service = personaService;
    }
    
    @GetMapping
    public String listarPersonas(Model model) {
        model.addAttribute("personas", this.service.listarTodas());
        return "personas";
    }
    
    @GetMapping("/nueva")
    public String mostrarFormularioCrear(Model model) {
        model.addAttribute("persona", new Persona());
        return "formulario";
    }
    
    @PostMapping
    public String guardarPersona(@ModelAttribute Persona persona) {
        this.service.guardar(persona);
        return "redirect:/personas";
    }
    
    @GetMapping("/editar/{id}")
    public String mostrarFormularioEditar(@PathVariable Long id, Model model) {
        model.addAttribute("persona", this.service.buscarPorId(id).orElseThrow(() -> new IllegalArgumentException("ID invalido" + id)));
        return "formulario";
    }
    
    @GetMapping("/reporte/pdf")
    public void generarPdf(HttpServletResponse response) throws IOException {
        response.setContentType("application/pdf");
        response.setHeader("Content-Disposition", "inline; filename=personas_reporte.pdf");
        
        PdfWriter write = new PdfWriter(response.getOutputStream());
        Document document = new Document(new com.itextpdf.kernel.pdf.PdfDocument(write));
        
        document.add(new Paragraph("Reporte de personas").setBold().setFontSize(18));
        
        Table table = new Table(3);
        table.addCell("ID");
        table.addCell("Nombre");
        table.addCell("Apellido");
        
        List<Persona> personas = this.service.listarTodas();
        for (Persona persona : personas) {
            table.addCell(persona.getId().toString());
            table.addCell(persona.getNombre());
            table.addCell(persona.getApellido());
        }
        
        document.add(table);
        document.close();
    }
    
    @GetMapping("/reporte/excel")
    public void generarExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=personas_reporte.xlsx");
        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Personas");
        
        Row headerRow = sheet.createRow(0);
        String[] columnHeaders = {"ID", "Nombre", "Apellidos"};
        for (int i = 0; i < columnHeaders.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columnHeaders[i]);
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            cell.setCellStyle(style);
        }
        
        List<Persona> personas = this.service.listarTodas();
        int rowIndex = 1;
        for (Persona persona : personas) {
            Row row = sheet.createRow(rowIndex++);
            row.createCell(0).setCellValue(persona.getId());
            row.createCell(1).setCellValue(persona.getNombre());
            row.createCell(2).setCellValue(persona.getApellido());
        }
        
        /*for (int i = 0; columnHeaders.length; i++) {
            sheet.autoSizeColumn(i);
        }*/
        
        workbook.write(response.getOutputStream());
        workbook.close();
        
    }
    
}
