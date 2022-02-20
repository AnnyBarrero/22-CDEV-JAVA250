package com.example.demo.service.export;
import com.example.demo.entity.Client;
import com.example.demo.repository.ClientRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import java.io.PrintWriter;
import java.util.List;
@Service
public class ClientExportCVSService {

    @Autowired
    private ClientRepository clientRepository;
    public void export(PrintWriter writer) {
        List<Client> clients = clientRepository.findAll();
        writer.println("Nom;Prénom");
        for (Client client : clients) {
            writer.println(client.getNom() + ";" + client.getPrenom());
        }
    }


}