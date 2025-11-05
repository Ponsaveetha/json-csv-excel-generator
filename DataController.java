package com.example.demo.controller;

import com.example.demo.service.DataService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.util.*;

@Controller
public class DataController {

    @Autowired
    private DataService dataService;

    private List<Map<String, Object>> currentData = new ArrayList<>();

    @GetMapping("/")
    public String index(Model model) {
        model.addAttribute("headers", currentData.isEmpty() ? List.of("Name", "Email") : new ArrayList<>(currentData.get(0).keySet()));
        model.addAttribute("data", currentData);
        return "index";
    }

   @PostMapping("/generate")
@ResponseBody
public List<Map<String, Object>> generateData(@RequestBody Map<String, Object> payload) {
    List<String> headers = (List<String>) payload.get("headers");
    int rowCount = (int) payload.getOrDefault("rowCount", 5);
    return dataService.generateTestData(headers, rowCount);
}


    @PostMapping("/import")
    public String importFile(@RequestParam("file") MultipartFile file, Model model) throws Exception {
        String fileName = file.getOriginalFilename();
        if (fileName == null) return "redirect:/";

        if (fileName.endsWith(".json")) {
            currentData = dataService.importFromJson(file);
        } else if (fileName.endsWith(".csv")) {
            currentData = dataService.importFromCsv(file);
        } else if (fileName.endsWith(".xlsx")) {
            currentData = dataService.importFromExcel(file);
        }

        model.addAttribute("headers", currentData.get(0).keySet());
        model.addAttribute("data", currentData);
        return "index";
    }

    @PostMapping("/save-edits")
    public String saveEdits(@RequestParam Map<String, String> formData, Model model) {
        // formData contains flattened data from editable table
        Set<String> headers = new LinkedHashSet<>();
        List<Map<String, Object>> newData = new ArrayList<>();

        int rowCount = Integer.parseInt(formData.get("rowCount"));
        int colCount = Integer.parseInt(formData.get("colCount"));

        // Collect headers
        for (int c = 0; c < colCount; c++) {
            headers.add(formData.get("header_" + c));
        }

        // Collect rows
        for (int r = 0; r < rowCount; r++) {
            Map<String, Object> row = new LinkedHashMap<>();
            for (int c = 0; c < colCount; c++) {
                String header = formData.get("header_" + c);
                String value = formData.get("cell_" + r + "_" + c);
                row.put(header, value);
            }
            newData.add(row);
        }

        currentData = newData;
        model.addAttribute("headers", headers);
        model.addAttribute("data", currentData);
        return "index";
    }

    // Export endpoints
    @GetMapping("/export/json")
    public ResponseEntity<InputStreamResource> exportJson() throws Exception {
        byte[] data = dataService.exportAsJson(currentData);
        return createResponse(data, "data.json", "application/json");
    }

    @GetMapping("/export/csv")
    public ResponseEntity<InputStreamResource> exportCsv() throws Exception {
        byte[] data = dataService.exportAsCsv(currentData);
        return createResponse(data, "data.csv", "text/csv");
    }

    @GetMapping("/export/excel")
    public ResponseEntity<InputStreamResource> exportExcel() throws Exception {
        byte[] data = dataService.exportAsExcel(currentData);
        return createResponse(data, "data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    private ResponseEntity<InputStreamResource> createResponse(byte[] data, String filename, String type) {
        InputStreamResource resource = new InputStreamResource(new ByteArrayInputStream(data));
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType(type))
                .body(resource);
    }
}
