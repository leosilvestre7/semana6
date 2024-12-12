/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.example.semana5.service;

import com.example.semana5.repository.PersonaRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import com.example.semana5.model.Persona;
import java.util.List;
import java.util.Optional;

/**
 *
 * @author leoru
 */
@Service
public class PersonaService {
    
    @Autowired
    private PersonaRepository repository;
    
    /**
     * Funcion para listar la tabla persona
     * @return 
     */
    public List<Persona> listarTodas() {
        return repository.findAll();
    }
    
    /**
     * Funcion para guardar datos de una persona
     * @param persona 
     */
    public void guardar(Persona persona) {
        repository.save(persona);
    }
    
    /**
     * Funcion para buscar una persona por id
     * @param id
     * @return 
     */
    public Optional<Persona> buscarPorId(Long id) {
        return repository.findById(id);
    }
    
    /**
     * Funcion para eliminar el registro de una persona
     * @param id 
     */
    public void eliminar(Long id) {
        repository.deleteById(id);
    }
    
    
}
