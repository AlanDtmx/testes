package com.prefeitura;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileWriter;
import java.io.InputStream;
import java.util.Locale;


@RestController
@RequestMapping("/upload")
public class UploadController {
    @PostMapping("/upload")
    public ResponseEntity<String> verificarArquivo(@RequestParam("file") MultipartFile file,
                                                   @RequestParam("evento50") String evento50,
                                                   @RequestParam("evento100") String evento100,
                                                   @RequestParam("eventoNoturno") String eventoNoturno,
                                                   @RequestParam("eventoAtrasos") String eventoAtrasos

    ) {

        try {
            if (file.isEmpty()) {
                return ResponseEntity.badRequest().body("Arquivo está vazio.");
            }

            InputStream inputStream = file.getInputStream();
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            StringBuilder resultado = new StringBuilder();

            for (int i = 6; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Coluna B = índice 1 (matrícula)
                Cell matriculaCell = row.getCell(1);
                if (matriculaCell == null || matriculaCell.getCellType() != CellType.NUMERIC) continue;

                String matricula = String.valueOf((long) matriculaCell.getNumericCellValue());

                // Coluna C = índice 2 (horas 50%)
                Cell horas50Cell = row.getCell(2);
                if (horas50Cell != null && horas50Cell.getCellType() == CellType.NUMERIC) {
                    String horas50 = String.format(Locale.forLanguageTag("pt-BR"), "%.2f", horas50Cell.getNumericCellValue()).replace('.', ',');
                    resultado.append(matricula).append(";").append(evento50).append(";0;")
                            .append(horas50).append(";1\n");
                }

                // Coluna D = índice 3 (horas 100%)
                Cell horas100Cell = row.getCell(3);
                if (horas100Cell != null && horas100Cell.getCellType() == CellType.NUMERIC) {
                    String horas100 = String.format(Locale.forLanguageTag("pt-BR"), "%.2f", horas100Cell.getNumericCellValue()).replace('.', ',');
                    resultado.append(matricula).append(";").append(evento100).append(";0;")
                            .append(horas100).append(";1\n");
                }

                // Coluna E = índice 4 (adicional noturno)
                Cell adicionalNoturnoCell = row.getCell(4);
                if (adicionalNoturnoCell != null && adicionalNoturnoCell.getCellType() == CellType.NUMERIC) {
                    String adicionalNoturno = String.format(Locale.forLanguageTag("pt-BR"), "%.2f", adicionalNoturnoCell.getNumericCellValue()).replace('.', ',');
                    resultado.append(matricula).append(";").append(eventoNoturno).append(";0;")
                            .append(adicionalNoturno).append(";1\n");
                }

                // Coluna F = índice 5 (atrasos)
                Cell atrasosCell = row.getCell(5);
                if (atrasosCell != null && atrasosCell.getCellType() == CellType.NUMERIC) {
                    String atrasos = String.format(Locale.forLanguageTag("pt-BR"), "%.2f", atrasosCell.getNumericCellValue()).replace('.', ',');
                    resultado.append(matricula).append(";").append(eventoAtrasos).append(";0;")
                            .append(atrasos).append(";1\n");
                }
            }

            workbook.close();
            File outputDir = new File("output");
            if (!outputDir.exists()) outputDir.mkdirs(); // cria a pasta se não existir

            File arquivo = new File(outputDir, "resultado.txt");
            FileWriter writer = new FileWriter(arquivo);
            writer.write(resultado.toString());
            writer.close();

            return ResponseEntity.ok("Arquivo gerado com sucesso e salvo em: " + arquivo.getAbsolutePath());
        } catch (Exception e) {
            return ResponseEntity.status(500).body("Erro ao abrir o arquivo: " + e.getMessage());
        }
    }

    }
