package com.example.demo.controller.clientsidetemplating;

import com.example.demo.dto.ClientDto;
import com.example.demo.service.ClientService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("rest/Clients")

public class ClientController {
    @Autowired
    private ClientService clientService;

    /**
     * Exposition d'une api déclenchée sur l'url http://..../clients.
     *
     * @return la liste des clients au format JSON. Voir la classe ClientDto pour le détail du format.
     */
    @GetMapping
    public List<ClientDto> getClients() {
        return clientService.findAll();
    }
}
