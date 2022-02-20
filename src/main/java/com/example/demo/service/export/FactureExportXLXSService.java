package com.example.demo.service.export;

import com.example.demo.entity.Client;
import com.example.demo.repository.ClientRepository;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.FactureRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

@Service
public class FactureExportXLXSService {

    @Autowired
    private FactureRepository factureRepository;
    private ClientRepository clientRepository;

    public void export(OutputStream outputStream) throws IOException {
        Workbook workbook = new XSSFWorkbook();

        List<Client> clients = clientRepository.findAll();

        //créer un objet style
        CellStyle cellStyleHeader = workbook.createCellStyle();
        Font fontHeader = workbook.createFont();
        fontHeader.setColor(IndexedColors.PINK.index);
        fontHeader.setBold(true);
        cellStyleHeader.setFont(fontHeader);
        applyBorder(cellStyleHeader);
        CellStyle cellStyleData = workbook.createCellStyle();
        applyBorder(cellStyleData);


        int isheet = 1;
        for (Client client : clients) {

            //List<Facture> factures= factureRepository.findAll();
            List<Facture> factures= factureRepository.findByClientId(client.getId());

            Sheet sheetC = workbook.createSheet(client.getNom() + client.getPrenom());
            Row rowH = sheetC.createRow(0);
            Cell cellHeader0 = rowH.createCell(0);
            cellHeader0.setCellValue("Surname");
            cellHeader0.setCellStyle(cellStyleHeader);
            Cell cellHeader1 = rowH.createCell(1);
            cellHeader1.setCellValue("Name");
            cellHeader1.setCellStyle(cellStyleHeader);
            Cell cellHeader2 = rowH.createCell(2);
            cellHeader2.setCellValue("Date of Birth");
            cellHeader2.setCellStyle(cellStyleHeader);
            Cell cellHeader3 = rowH.createCell(3);
            cellHeader3.setCellValue("Number of factures");
            cellHeader3.setCellStyle(cellStyleHeader);


            Row row = sheetC.createRow(1);
            Cell cell0 = row.createCell(0);
            cell0.setCellValue(client.getNom());
            Cell cell1 = row.createCell(1);
            cell1.setCellValue(client.getPrenom());
            Cell cell2 = row.createCell(2);
            cell2.setCellValue(client.getDateNais().getYear());
            Cell cell3 = row.createCell(3);
            cell3.setCellValue(factures.stream().count());

            for (Facture facture : factures) {
                Sheet sheetF = workbook.createSheet("Facture N°" + facture.getId() + "-" + isheet);
                Row rowHf = sheetF.createRow(0);
                Cell cellHeaderf0 = rowHf.createCell(0);
                cellHeaderf0.setCellValue("Article");
                cellHeaderf0.setCellStyle(cellStyleHeader);
                Cell cellHeaderf1 = rowHf.createCell(1);
                cellHeaderf1.setCellValue("Quantity");
                cellHeaderf1.setCellStyle(cellStyleHeader);
                Cell cellHeaderf2 = rowHf.createCell(2);
                cellHeaderf2.setCellValue("Price");
                cellHeaderf2.setCellStyle(cellStyleHeader);
                Cell cellHeaderf3 = rowHf.createCell(3);
                cellHeaderf3.setCellValue("Total price");
                cellHeaderf3.setCellStyle(cellStyleHeader);
                Cell cellHeaderf4 = rowHf.createCell(4);
                cellHeaderf4.setCellValue(" ");
                int rowfc = 1;
                Double total = 0.0;
                for (LigneFacture ligneFacture : facture.getLigneFactures()) {

                    Row rowf = sheetF.createRow(rowfc);
                    Cell cellf0 = rowf.createCell(0);
                    cellf0.setCellValue(ligneFacture.getArticle().getLibelle());
                    Cell cellf1 = rowf.createCell(1);
                    cellf1.setCellValue(ligneFacture.getQuantite());
                    Cell cellf2 = rowf.createCell(2);
                    cellf2.setCellValue(ligneFacture.getArticle().getPrix());
                    Cell cellf3 = rowf.createCell(3);
                    cellf3.setCellValue(ligneFacture.getArticle().getPrix() * ligneFacture.getQuantite());

                    total = total + (ligneFacture.getArticle().getPrix() * ligneFacture.getQuantite());
                    rowfc++;

                }

                Row rowTotal = sheetF.createRow(rowfc);
                Cell totally = rowTotal.createCell(0);
                totally.setCellValue("total");
                totally.setCellStyle(cellStyleHeader);
                Cell cellTotalPrice = rowTotal.createCell(3);
                cellTotalPrice.setCellValue(total);

                sheetF.autoSizeColumn(0);
                sheetF.autoSizeColumn(1);
                sheetF.autoSizeColumn(2);
                sheetF.autoSizeColumn(3);


                isheet++;

            }


        }
    }

    private void applyBorder(CellStyle cellStyleHeader) {
        cellStyleHeader.setBorderTop(BorderStyle.THICK);
        cellStyleHeader.setBorderRight(BorderStyle.THICK);
        cellStyleHeader.setBorderBottom(BorderStyle.THICK);
        cellStyleHeader.setBorderLeft(BorderStyle.THICK);
        cellStyleHeader.setTopBorderColor(IndexedColors.BLUE.index);
        cellStyleHeader.setRightBorderColor(IndexedColors.BLUE.index);
        cellStyleHeader.setBottomBorderColor(IndexedColors.BLUE.index);
        cellStyleHeader.setLeftBorderColor(IndexedColors.BLUE.index);
    }


}
