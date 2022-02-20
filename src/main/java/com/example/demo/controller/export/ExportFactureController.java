package com.example.demo.controller.export;
import com.example.demo.service.export.FactureExportPdfService;
import com.itextpdf.text.DocumentException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;

@Controller
@RequestMapping("export/factures")
public class ExportFactureController {
    @Autowired
    private FactureExportPdfService factureExportPdfService;

    //POUR DEMAIN
    @GetMapping("{id}/pdf")
    public void exportPdf(
            @PathVariable Long id,
            HttpServletRequest request,
            HttpServletResponse response
    ) throws IOException, DocumentException {
        response.setHeader("Content-Disposition", "attachment; filename=\"export-facture-" + id + ".pdf\"");
        OutputStream outputStream = response.getOutputStream();
        factureExportPdfService.export(outputStream, id);

    }
}
