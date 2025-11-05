package com.example.demo.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

/**
 * Enhanced DataService for generating, editing, importing, and exporting test data.
 * Supports JSON, CSV, and Excel formats.
 */
@Service
public class DataService {

    /**
     * Generate random test data with provided headers.
     */
    public List<Map<String, Object>> generateTestData(List<String> headers, int rowCount) {
        List<Map<String, Object>> dataList = new ArrayList<>();
        Random random = new Random();

        for (int i = 0; i < rowCount; i++) {
            Map<String, Object> row = new LinkedHashMap<>();
            for (String header : headers) {
                row.put(header, "Sample_" + (random.nextInt(900) + 100));
            }
            dataList.add(row);
        }
        return dataList;
    }

    // =====================================================================
    //                          JSON EXPORT / IMPORT
    // =====================================================================

    public byte[] exportAsJson(List<Map<String, Object>> data) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        return mapper.writerWithDefaultPrettyPrinter().writeValueAsBytes(data);
    }

    public List<Map<String, Object>> importFromJson(MultipartFile file) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        List<Map<String, Object>> list = Arrays.asList(mapper.readValue(file.getInputStream(), Map[].class));
        return normalizeData(list);
    }

    // =====================================================================
    //                          CSV EXPORT / IMPORT
    // =====================================================================

    public byte[] exportAsCsv(List<Map<String, Object>> data) throws IOException {
        if (data == null || data.isEmpty()) return new byte[0];

        StringBuilder sb = new StringBuilder();
        Set<String> headers = data.get(0).keySet();
        sb.append(String.join(",", headers)).append("\n");

        for (Map<String, Object> row : data) {
            List<String> values = new ArrayList<>();
            for (String header : headers) {
                Object value = row.get(header);
                String safeValue = value != null ? value.toString().replace("\"", "\"\"") : "";
                if (safeValue.contains(",") || safeValue.contains("\"")) {
                    safeValue = "\"" + safeValue + "\"";
                }
                values.add(safeValue);
            }
            sb.append(String.join(",", values)).append("\n");
        }

        return sb.toString().getBytes();
    }

    public List<Map<String, Object>> importFromCsv(MultipartFile file) throws IOException {
        List<Map<String, Object>> dataList = new ArrayList<>();

        try (BufferedReader br = new BufferedReader(new InputStreamReader(file.getInputStream()))) {
            String headerLine = br.readLine();
            if (headerLine == null) return Collections.emptyList();

            String[] headers = headerLine.split(",");
            String line;

            while ((line = br.readLine()) != null) {
                String[] values = parseCsvLine(line);
                Map<String, Object> row = new LinkedHashMap<>();
                for (int i = 0; i < headers.length; i++) {
                    row.put(headers[i].trim(), values.length > i ? values[i].trim() : "");
                }
                dataList.add(row);
            }
        }

        return normalizeData(dataList);
    }

    // CSV parsing helper
    private String[] parseCsvLine(String line) {
        List<String> result = new ArrayList<>();
        StringBuilder sb = new StringBuilder();
        boolean inQuotes = false;

        for (char c : line.toCharArray()) {
            if (c == '\"') {
                inQuotes = !inQuotes;
            } else if (c == ',' && !inQuotes) {
                result.add(sb.toString());
                sb.setLength(0);
            } else {
                sb.append(c);
            }
        }
        result.add(sb.toString());
        return result.toArray(new String[0]);
    }

    // =====================================================================
    //                          EXCEL EXPORT / IMPORT
    // =====================================================================

    public byte[] exportAsExcel(List<Map<String, Object>> data) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("TestData");

        if (!data.isEmpty()) {
            Row headerRow = sheet.createRow(0);
            int colIndex = 0;
            for (String key : data.get(0).keySet()) {
                Cell cell = headerRow.createCell(colIndex++);
                cell.setCellValue(key);
                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                style.setFont(font);
                cell.setCellStyle(style);
            }

            int rowIndex = 1;
            for (Map<String, Object> rowData : data) {
                Row row = sheet.createRow(rowIndex++);
                int i = 0;
                for (Object value : rowData.values()) {
                    Cell cell = row.createCell(i++);
                    cell.setCellValue(value != null ? value.toString() : "");
                }
            }

            for (int i = 0; i < data.get(0).keySet().size(); i++) {
                sheet.autoSizeColumn(i);
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        return out.toByteArray();
    }

    public List<Map<String, Object>> importFromExcel(MultipartFile file) throws IOException {
        List<Map<String, Object>> dataList = new ArrayList<>();

        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rows = sheet.iterator();

        List<String> headers = new ArrayList<>();
        if (rows.hasNext()) {
            Row headerRow = rows.next();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }
        }

        while (rows.hasNext()) {
            Row row = rows.next();
            Map<String, Object> data = new LinkedHashMap<>();
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = row.getCell(i);
                data.put(headers.get(i), getCellValue(cell));
            }
            dataList.add(data);
        }

        workbook.close();
        return normalizeData(dataList);
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            default -> "";
        };
    }

    // =====================================================================
    //                      DYNAMIC HEADER / EDIT SUPPORT
    // =====================================================================

    /**
     * Add a new header (column) to all rows dynamically.
     * New column will be initialized with an empty string.
     */
    public List<Map<String, Object>> addHeader(List<Map<String, Object>> data, String newHeader) {
        for (Map<String, Object> row : data) {
            row.putIfAbsent(newHeader, "");
        }
        return data;
    }

    /**
     * Update a single cell in the data list.
     */
    public List<Map<String, Object>> updateCell(List<Map<String, Object>> data, int rowIndex, String header, String newValue) {
        if (rowIndex >= 0 && rowIndex < data.size()) {
            Map<String, Object> row = data.get(rowIndex);
            if (row.containsKey(header)) {
                row.put(header, newValue);
            }
        }
        return data;
    }

    /**
     * Normalize data to ensure consistent header order and trimming.
     */
    private List<Map<String, Object>> normalizeData(List<Map<String, Object>> dataList) {
        if (dataList.isEmpty()) return dataList;

        Set<String> headers = new LinkedHashSet<>(dataList.get(0).keySet());
        List<Map<String, Object>> normalized = new ArrayList<>();

        for (Map<String, Object> row : dataList) {
            Map<String, Object> clean = new LinkedHashMap<>();
            for (String header : headers) {
                clean.put(header.trim(), row.getOrDefault(header, "").toString().trim());
            }
            normalized.add(clean);
        }

        return normalized;
    }
}
