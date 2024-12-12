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
    
}