package com.fi.taxfetcherapp;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.Font;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class TaxFetcherApp extends JFrame {
    private static final String OUTPUT_FOLDER = "output";
    private static final String BEXAR_BASE_URL = "https://bexar.acttax.com/act_webdev/bexar/showdetail2.jsp?can=";

    private JComboBox<String> countyComboBox;
    private JButton uploadButton;
    private JButton fetchButton;
    private JProgressBar progressBar;
    private JTextArea logArea;
    private JFileChooser fileChooser;
    private File selectedFile;
    private JPanel mainPanel;
    private ExecutorService executorService;
    private JLabel fileLabel; // Added instance variable for fileLabel

    public TaxFetcherApp() {
        initializeComponents();
        setupUI();
        setupEventHandlers();
        executorService = Executors.newFixedThreadPool(5);
    }

    private void initializeComponents() {
        countyComboBox = new JComboBox<>(new String[]{"Bexar", "Dallas"});
        uploadButton = new JButton("Upload Excel File");
        fetchButton = new JButton("Fetch Tax Details");
        fetchButton.setEnabled(false);
        progressBar = new JProgressBar(0, 100);
        progressBar.setIndeterminate(true);
        progressBar.setVisible(false);

        logArea = new JTextArea(10, 50);
        logArea.setEditable(false);
        logArea.setFont(new java.awt.Font("Monospaced", java.awt.Font.PLAIN, 12));

        fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));

        mainPanel = new JPanel(new BorderLayout(10, 10));
        fileLabel = new JLabel("Selected File: None"); // Initialize fileLabel
    }

    private void setupUI() {
        setTitle("Professional Tax Fetcher - Bexar & Dallas County");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setMinimumSize(new java.awt.Dimension(900, 700));
        setPreferredSize(new java.awt.Dimension(1000, 800));

        // Create title panel
        JPanel titlePanel = createTitlePanel();

        // Create main content panel
        JPanel contentPanel = new JPanel(new BorderLayout(15, 15));
        contentPanel.setBorder(new EmptyBorder(20, 20, 20, 20));

        // File selection panel
        JPanel filePanel = createFileSelectionPanel();

        // Progress panel
        JPanel progressPanel = createProgressPanel();

        // Log panel
        JPanel logPanel = createLogPanel();

        contentPanel.add(filePanel, BorderLayout.NORTH);
        contentPanel.add(progressPanel, BorderLayout.CENTER);
        contentPanel.add(logPanel, BorderLayout.SOUTH);

        mainPanel.add(titlePanel, BorderLayout.NORTH);
        mainPanel.add(contentPanel, BorderLayout.CENTER);

        add(mainPanel);

        // Center the window
        setLocationRelativeTo(null);
        pack();
    }

    private JPanel createTitlePanel() {
        JPanel titlePanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        titlePanel.setBorder(BorderFactory.createEmptyBorder(10, 0, 20, 0));

        JLabel titleLabel = new JLabel("Professional Tax Fetcher");
        titleLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 24));
        titleLabel.setForeground(new java.awt.Color(45, 55, 72));

        JLabel subtitleLabel = new JLabel("Bexar & Dallas County Property Tax Information");
        subtitleLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 14));
        subtitleLabel.setForeground(new java.awt.Color(100, 100, 100));

        titlePanel.add(titleLabel);
        titlePanel.add(new JLabel("  "));
        titlePanel.add(subtitleLabel);

        return titlePanel;
    }

    private JPanel createFileSelectionPanel() {
        JPanel filePanel = new JPanel(new BorderLayout(10, 10));
        filePanel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new java.awt.Color(200, 200, 200)),
                "File Selection"
        ));

        JPanel controlsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));

        JLabel countyLabel = new JLabel("County:");
        countyLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));

        fileLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        fileLabel.setForeground(java.awt.Color.GRAY);

        controlsPanel.add(countyLabel);
        controlsPanel.add(countyComboBox);
        controlsPanel.add(new JLabel("  "));
        controlsPanel.add(uploadButton);
        controlsPanel.add(new JLabel("  "));
        controlsPanel.add(fetchButton);
        controlsPanel.add(Box.createHorizontalGlue());
        controlsPanel.add(new JLabel("Status: "));
        controlsPanel.add(fileLabel);

        filePanel.add(controlsPanel, BorderLayout.CENTER);
        return filePanel;
    }

    private JPanel createProgressPanel() {
        JPanel progressPanel = new JPanel(new BorderLayout(10, 10));
        progressPanel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new java.awt.Color(200, 200, 200)),
                "Processing Status"
        ));

        progressPanel.add(progressBar, BorderLayout.CENTER);
        return progressPanel;
    }

    private JPanel createLogPanel() {
        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new java.awt.Color(200, 200, 200)),
                "Activity Log"
        ));

        JScrollPane scrollPane = new JScrollPane(logArea);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        logPanel.add(scrollPane, BorderLayout.CENTER);
        return logPanel;
    }

    private void setupEventHandlers() {
        uploadButton.addActionListener(e -> handleFileUpload());
        fetchButton.addActionListener(e -> handleFetchRequest());
    }

    private void handleFileUpload() {
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            selectedFile = fileChooser.getSelectedFile();
            fileLabel.setText("Selected File: " + selectedFile.getName());
            fileLabel.setForeground(java.awt.Color.BLACK);
            fetchButton.setEnabled(true);
            logMessage("File selected: " + selectedFile.getAbsolutePath());
        }
    }

    private void handleFetchRequest() {
        if (selectedFile == null) {
            JOptionPane.showMessageDialog(this, "Please select a file first.",
                    "No File Selected", JOptionPane.WARNING_MESSAGE);
            return;
        }

        String selectedCounty = (String) countyComboBox.getSelectedItem();
        if (!"Bexar".equals(selectedCounty)) {
            JOptionPane.showMessageDialog(this, "Currently only Bexar County is supported.",
                    "County Not Supported", JOptionPane.WARNING_MESSAGE);
            return;
        }

        new Thread(() -> processFile(selectedFile)).start();
    }

    private void processFile(File inputFile) {
        try {
            logMessage("Starting processing of: " + inputFile.getName());
            progressBar.setVisible(true);
            progressBar.setString("Loading Excel file...");

            // Read Excel file
            List<Map<String, String>> records = readExcelFile(inputFile);
            logMessage("Loaded " + records.size() + " records from Excel file");

            // Create output directory
            Files.createDirectories(Paths.get(OUTPUT_FOLDER));

            // Process each record
            int processed = 0;
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Sheet1");
            createOutputHeader(outputSheet);

            for (Map<String, String> record : records) {
                String taxId = record.get("TAXID");
                if (taxId != null && !taxId.trim().isEmpty()) {
                    String accountNumber = extractAccountNumber(taxId);
                    if (accountNumber != null) {
                        try {
                            Map<String, String> updatedRecord = fetchTaxDetails(accountNumber, record);
                            writeRecordToSheet(outputSheet, updatedRecord);
                            processed++;
                            logMessage("Processed record " + processed + "/" + records.size() +
                                    " - Account: " + accountNumber);

                            // Update progress
                            int progress = (int) ((double) processed / records.size() * 100);
                            progressBar.setValue(progress);
                            progressBar.setString(progress + "% - Processed " + processed + " records");

                            // Small delay to be respectful to the server
                            Thread.sleep(500);
                        } catch (Exception e) {
                            logMessage("Error processing account " + accountNumber + ": " + e.getMessage());
                        }
                    }
                }
            }

            // Save output file
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String outputFileName = OUTPUT_FOLDER + "/Bexar_Tax_Details_" + timestamp + ".xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputFileName)) {
                outputWorkbook.write(fos);
            }

            // Store final values for use in lambda
            final int finalProcessed = processed;
            final String finalOutputFileName = outputFileName;
            final int finalRecordsSize = records.size();

            progressBar.setValue(100);
            progressBar.setString("Completed! Output saved to: " + finalOutputFileName);
            logMessage("Processing completed! Output file saved: " + finalOutputFileName);
            logMessage("Total records processed: " + finalProcessed + "/" + finalRecordsSize);

            // Show completion dialog
            SwingUtilities.invokeLater(() -> {
                JOptionPane.showMessageDialog(TaxFetcherApp.this,
                        "Processing completed successfully!\n" +
                                "Output file: " + finalOutputFileName + "\n" +
                                "Records processed: " + finalProcessed + "/" + finalRecordsSize,
                        "Process Complete", JOptionPane.INFORMATION_MESSAGE);
            });

        } catch (Exception e) {
            logMessage("Error processing file: " + e.getMessage());
            e.printStackTrace();
            SwingUtilities.invokeLater(() -> {
                JOptionPane.showMessageDialog(TaxFetcherApp.this,
                        "Error processing file: " + e.getMessage(),
                        "Processing Error", JOptionPane.ERROR_MESSAGE);
            });
        } finally {
            progressBar.setVisible(false);
            progressBar.setValue(0);
        }
    }

    private List<Map<String, String>> readExcelFile(File file) throws IOException, InvalidFormatException {
        List<Map<String, String>> records = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            // Get header row
            Row headerRow = sheet.getRow(0);
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(getCellValueAsString(cell).trim());
            }

            // Process data rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Map<String, String> record = new HashMap<>();
                    for (int j = 0; j < headers.size(); j++) {
                        Cell cell = row.getCell(j);
                        String value = getCellValueAsString(cell);
                        record.put(headers.get(j), value.trim());
                    }
                    records.add(record);
                }
            }
        }

        return records;
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((long) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private String extractAccountNumber(String taxId) {
        if (taxId.startsWith("ACCT")) {
            return taxId.substring(4).trim();
        }
        return taxId.trim();
    }

    private Map<String, String> fetchTaxDetails(String accountNumber, Map<String, String> originalRecord) {
        Map<String, String> updatedRecord = new HashMap<>(originalRecord);
        updatedRecord.put("ACCOUNT_NUMBER", accountNumber);

        try {
            String url = BEXAR_BASE_URL + accountNumber;
            logMessage("Fetching details for: " + url);

            // Set up connection with reasonable timeout
            HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
            connection.setRequestProperty("User-Agent",
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
            connection.setConnectTimeout(10000);
            connection.setReadTimeout(15000);

            Document doc = Jsoup.parse(connection.getInputStream(), "UTF-8", url);

            // Extract key information
            extractOwnerInfo(doc, updatedRecord);
            extractPropertyInfo(doc, updatedRecord);
            extractTaxInfo(doc, updatedRecord);
            extractPaymentInfo(doc, updatedRecord);
            extractValueInfo(doc, updatedRecord);

            updatedRecord.put("FETCH_STATUS", "SUCCESS");
        } catch (Exception e) {
            logMessage("Failed to fetch details for " + accountNumber + ": " + e.getMessage());
            updatedRecord.put("FETCH_STATUS", "ERROR: " + e.getMessage());
        }

        return updatedRecord;
    }

    private void extractOwnerInfo(Document doc, Map<String, String> record) {
        try {
            Elements ownerElements = doc.select("td:contains(Owner Name), td:contains(Owner)");
            if (!ownerElements.isEmpty()) {
                String ownerText = ownerElements.first().parent().text();
                record.put("CURRENT_OWNER_NAME", cleanText(ownerText));

                // Extract address components
                String fullAddress = extractAddressFromText(ownerText);
                if (fullAddress != null) {
                    Map<String, String> addressParts = parseAddress(fullAddress);
                    record.putAll(addressParts);
                }
            }
        } catch (Exception e) {
            record.put("CURRENT_OWNER_NAME", "Unable to extract");
        }
    }

    private void extractPropertyInfo(Document doc, Map<String, String> record) {
        try {
            Elements propElements = doc.select("td:contains(Property Address), td:contains(Property)");
            if (!propElements.isEmpty()) {
                String propText = propElements.first().parent().text();
                record.put("CURRENT_PROP_ADDRESS", cleanText(propText));

                String fullPropAddress = extractAddressFromText(propText);
                if (fullPropAddress != null) {
                    Map<String, String> propAddressParts = parseAddress(fullPropAddress);
                    record.put("CURRENT_PROP_STREET", propAddressParts.getOrDefault("street", ""));
                    record.put("CURRENT_PROP_CITY", propAddressParts.getOrDefault("city", ""));
                    record.put("CURRENT_PROP_STATE", propAddressParts.getOrDefault("state", ""));
                    record.put("CURRENT_PROP_ZIP", propAddressParts.getOrDefault("zip", ""));
                }
            }
        } catch (Exception e) {
            record.put("CURRENT_PROP_ADDRESS", "Unable to extract");
        }
    }

    private void extractTaxInfo(Document doc, Map<String, String> record) {
        try {
            Elements taxElements = doc.select("td:contains($), span:contains($)");
            for (Element element : taxElements) {
                String text = element.text().trim();
                if (text.matches(".*\\$\\d+[,\\d]*\\.\\d{2}.*")) {
                    String taxAmount = text.replaceAll("[^0-9.]", "").replaceFirst("^0+", "");
                    if (!taxAmount.isEmpty()) {
                        record.put("CURRENT_TAX_DUE", "$" + taxAmount);
                        break;
                    }
                }
            }
        } catch (Exception e) {
            record.put("CURRENT_TAX_DUE", "Unable to extract");
        }
    }

    private void extractPaymentInfo(Document doc, Map<String, String> record) {
        try {
            Elements paymentElements = doc.select("td:contains(Payment), td:contains(Last Payment)");
            if (!paymentElements.isEmpty()) {
                String paymentText = paymentElements.first().parent().text();
                record.put("LAST_PAYMENT_INFO", cleanText(paymentText));
            }
        } catch (Exception e) {
            record.put("LAST_PAYMENT_INFO", "Unable to extract");
        }
    }

    private void extractValueInfo(Document doc, Map<String, String> record) {
        try {
            Elements valueElements = doc.select("td:contains(Value), td:contains(Assessed)");
            for (Element element : valueElements) {
                String text = element.text().trim();
                if (text.contains("$") && text.matches(".*\\$\\d+[,\\d]*.*")) {
                    String valueType = cleanText(element.previousElementSibling().text());
                    String valueAmount = text.replaceAll("[^0-9,]", "").replace(",", "");

                    if (valueType.toLowerCase().contains("land")) {
                        record.put("CURRENT_LAND_VALUE", "$" + valueAmount);
                    } else if (valueType.toLowerCase().contains("improvement") ||
                            valueType.toLowerCase().contains("improved")) {
                        record.put("CURRENT_IMPROVEMENT_VALUE", "$" + valueAmount);
                    } else if (valueType.toLowerCase().contains("total") ||
                            valueType.toLowerCase().contains("market")) {
                        record.put("CURRENT_TOTAL_VALUE", "$" + valueAmount);
                    }
                }
            }
        } catch (Exception e) {
            record.put("CURRENT_TOTAL_VALUE", "Unable to extract");
        }
    }

    private String cleanText(String text) {
        if (text == null) return "";
        return text.replaceAll("\\s+", " ").trim();
    }

    private String extractAddressFromText(String text) {
        String[] patterns = {
                "(\\d+\\s+[A-Za-z\\s]+(?:\\s+(?:St|Street|Ave|Avenue|Dr|Drive|Rd|Road|Ln|Lane|Blvd|Boulevard|Ct|Court|Pl|Place|Way)))[,\\s]+([A-Za-z\\s]+),\\s*([A-Z]{2})\\s*(\\d{5})",
                "(\\d+\\s+[A-Za-z\\s]+)[,\\s]+([A-Za-z\\s]+),\\s*([A-Z]{2})\\s*(\\d{5})",
                "([A-Za-z\\s]+(?:\\s+(?:St|Ave|Dr|Rd|Ln|Blvd|Ct|Pl|Way)))[,\\s]+([A-Za-z\\s]+),\\s*([A-Z]{2})\\s*(\\d{5})"
        };

        for (String pattern : patterns) {
            java.util.regex.Pattern p = java.util.regex.Pattern.compile(pattern);
            java.util.regex.Matcher m = p.matcher(text);
            if (m.find()) {
                return m.group(0);
            }
        }
        return null;
    }

    private Map<String, String> parseAddress(String address) {
        Map<String, String> parts = new HashMap<>();

        String[] components = address.split(",");
        if (components.length >= 2) {
            String street = components[0].trim();
            String cityStateZip = components[1].trim();

            parts.put("street", street);

            String[] cityParts = cityStateZip.split("\\s+");
            if (cityParts.length >= 3) {
                StringBuilder city = new StringBuilder();
                for (int i = 0; i < cityParts.length - 2; i++) {
                    if (i > 0) city.append(" ");
                    city.append(cityParts[i]);
                }
                parts.put("city", city.toString().trim());
                parts.put("state", cityParts[cityParts.length - 2]);
                parts.put("zip", cityParts[cityParts.length - 1]);
            }
        }

        return parts;
    }

    private void createOutputHeader(Sheet sheet) {
        String[] headers = {
                "JDX", "ACCOUNT_NUMBER", "PropID", "LastRun", "OwnerName", "OwnerStreet",
                "OwnerCity", "OwnerState", "OwnerZIP", "PropStreet", "PropCity",
                "PropState", "PropZIP", "Description", "Exemptions", "Lawsuit", "BK",
                "Tax", "Fees", "PriorDue", "LastPayment", "LastPaymentDate",
                "LastPayer", "PendingPayment", "PendingPaymentDate", "ValueAss",
                "ValueLand", "ValueImp", "CurrentDue", "TOTAL DUE", "LTV",
                "law suit active", "FEES2", "TT W FEES 4 PMT", "RATE", "APR",
                "pmt", "best payment option", "Back of card repayment obligation",
                "Obligation IF you use entire term", "lesser obligation", "IF PAID BY",
                "ESTIMATED MAX PURCHASE PRICE", "CASH TO CUSTOMER", "WIGGLE ROOM",
                "FORECLOSURE", "tax loan amount", "lender name", "MAILER/DELTE",
                "pmt 24 mts", "MobileHome", "UNIQUE",
                "CURRENT_OWNER_NAME", "CURRENT_OWNER_STREET", "CURRENT_OWNER_CITY",
                "CURRENT_OWNER_STATE", "CURRENT_OWNER_ZIP", "CURRENT_PROP_ADDRESS",
                "CURRENT_PROP_STREET", "CURRENT_PROP_CITY", "CURRENT_PROP_STATE",
                "CURRENT_PROP_ZIP", "CURRENT_TAX_DUE", "LAST_PAYMENT_INFO",
                "CURRENT_TOTAL_VALUE", "CURRENT_LAND_VALUE", "CURRENT_IMPROVEMENT_VALUE",
                "FETCH_STATUS", "FETCH_DATE"
        };

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);

            // Style the header
            CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
            org.apache.poi.ss.usermodel.Font headerFont = sheet.getWorkbook().createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 11);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            cell.setCellStyle(headerStyle);
        }

        // Auto-size columns
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void writeRecordToSheet(Sheet sheet, Map<String, String> record) {
        int lastRowNum = sheet.getLastRowNum() + 1;
        Row row = sheet.createRow(lastRowNum);

        String[] headers = {
                "JDX", "ACCOUNT_NUMBER", "PropID", "LastRun", "OwnerName", "OwnerStreet",
                "OwnerCity", "OwnerState", "OwnerZIP", "PropStreet", "PropCity",
                "PropState", "PropZIP", "Description", "Exemptions", "Lawsuit", "BK",
                "Tax", "Fees", "PriorDue", "LastPayment", "LastPaymentDate",
                "LastPayer", "PendingPayment", "PendingPaymentDate", "ValueAss",
                "ValueLand", "ValueImp", "CurrentDue", "TOTAL DUE", "LTV",
                "law suit active", "FEES2", "TT W FEES 4 PMT", "RATE", "APR",
                "pmt", "best payment option", "Back of card repayment obligation",
                "Obligation IF you use entire term", "lesser obligation", "IF PAID BY",
                "ESTIMATED MAX PURCHASE PRICE", "CASH TO CUSTOMER", "WIGGLE ROOM",
                "FORECLOSURE", "tax loan amount", "lender name", "MAILER/DELTE",
                "pmt 24 mts", "MobileHome", "UNIQUE",
                "CURRENT_OWNER_NAME", "CURRENT_OWNER_STREET", "CURRENT_OWNER_CITY",
                "CURRENT_OWNER_STATE", "CURRENT_OWNER_ZIP", "CURRENT_PROP_ADDRESS",
                "CURRENT_PROP_STREET", "CURRENT_PROP_CITY", "CURRENT_PROP_STATE",
                "CURRENT_PROP_ZIP", "CURRENT_TAX_DUE", "LAST_PAYMENT_INFO",
                "CURRENT_TOTAL_VALUE", "CURRENT_LAND_VALUE", "CURRENT_IMPROVEMENT_VALUE",
                "FETCH_STATUS", "FETCH_DATE"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            String value = record.getOrDefault(headers[i], "");
            cell.setCellValue(value.isEmpty() ? "" : value);
        }

        // Add fetch timestamp
        Cell timestampCell = row.createCell(headers.length - 1);
        timestampCell.setCellValue(LocalDateTime.now().toString());
    }

    private void logMessage(String message) {
        SwingUtilities.invokeLater(() -> {
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"));
            logArea.append("[" + timestamp + "] " + message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }

            TaxFetcherApp app = new TaxFetcherApp();
            app.setVisible(true);
        });
    }
}