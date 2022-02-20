package com.example.demo.service.export;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.ClientRepository;
import com.example.demo.repository.FactureRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.util.List;

@Service
public class ClientExportCVSService {
    @Autowired
    private ClientRepository clientRepository;

    public void export(PrintWriter writer) {
        List<Client> clients = clientRepository.findAll();
        writer.println("Nom;Prénom;Age");
        for (Client client : clients) {
            int age = LocalDate.now().getYear()-client.getDateNais().getYear();
            writer.println(client.getNom() + ";" + client.getPrenom()+";"+age);
        }
    }
}

