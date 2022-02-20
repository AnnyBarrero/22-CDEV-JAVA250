package com.example.demo.entity;

import javax.persistence.*;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import java.time.LocalDate;
import java.util.Date;


/**
 * Entity représentant un client.
 */
@Entity
public class Client {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @Column
    private String nom;

    @Column
    private String prenom;

    @Column
    private LocalDate dateNais;

    public Client(String nom, String prenom, LocalDate dateNais) {
        this.nom = nom;
        this.prenom = prenom;
        this.dateNais= dateNais;
    }

    public Client() {

    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getNom() {
        return nom;
    }

    public void setNom(String nom) {
        this.nom = nom;
    }

    public String getPrenom() {
        return prenom;
    }

    public void setPrenom(String prenom) {
        this.prenom = prenom;
    }

    public LocalDate getDateNais() {
        return dateNais;
    }

    public void setDateNais(LocalDate dateNais) {
        this.dateNais = dateNais;
    }
}
