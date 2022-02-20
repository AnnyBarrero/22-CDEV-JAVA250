package com.example.demo.service;

import com.example.demo.dto.ClientDto;
import com.example.demo.entity.Client;

import java.util.List;

public interface ClientService {
    List<ClientDto> findAll();
}