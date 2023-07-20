package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class ExcelFileReader extends JFrame {

    private final JLabel excelFileLabel;
    private final JTextArea logTextArea;
    private final JFileChooser fileChooserDialog;

    public ExcelFileReader() {
        setTitle("Excel Import");
        setSize(500, 400);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        excelFileLabel = new JLabel("Excel File:");
        JButton btnBrowse = new JButton("Browse");
        JButton btnImport = new JButton("Upload");
        logTextArea = new JTextArea();
        fileChooserDialog = new JFileChooser();

        JPanel topPanel = new JPanel();
        topPanel.add(excelFileLabel);
        topPanel.add(btnBrowse);
        add(topPanel, BorderLayout.NORTH);
        add(new JScrollPane(logTextArea), BorderLayout.CENTER);
        add(btnImport, BorderLayout.SOUTH);

        btnBrowse.addActionListener(e -> {
            int returnVal = fileChooserDialog.showOpenDialog(ExcelFileReader.this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fileChooserDialog.getSelectedFile();
                excelFileLabel.setText("Excel File: " + file.getName());
            }
        });

        btnImport.addActionListener(e -> {
            File selectedFile = fileChooserDialog.getSelectedFile();
            if (selectedFile != null) {
                importData(selectedFile);
            } else {
                JOptionPane.showMessageDialog(ExcelFileReader.this, "Please select an Excel file.");
            }
        });
    }
    private void importData(File selectedFile) {
        try {
            FileInputStream fis = new FileInputStream(selectedFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            List<List<String>> excelRows = new ArrayList<>(); // ArrayList to store the rows

            int excelDuplicateCount = 0;
            int dbDuplicateCount = 0;

            // Iterate over rows
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                // Get the cell value of the first column (employee name)
                Cell nameCell = row.getCell(1);
                String name = getCellValueAsString(nameCell);

                // Check if the name already exists in the database
                List<String> rowsFromDb = fetchDataFromDatabase();
                if (rowsFromDb.contains(name)) {
                    dbDuplicateCount++;
                    continue;
                }

                // Check if the name already exists in the excelRows list
                boolean isDuplicate = false;
                for (List<String> existingRow : excelRows) {
                    if (existingRow.get(0).equals(name)) {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate) {
                    List<String> rowData = new ArrayList<>(); // Move the rowData list inside the loop

                    // Iterate over cells in the row, starting from the second column (index 1)
                    for (int cellIndex = 1; cellIndex < row.getLastCellNum(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        // Get the cell value as a String
                        String cellValue = getCellValueAsString(cell);
                        rowData.add(cellValue);
                    }

                    excelRows.add(rowData);
                } else {
                    excelDuplicateCount++;
                }
            }

            // Insert non-duplicate rows into the database
            storeNonDuplicateRowsInDatabase(excelRows);

            List<String> excelNames = new ArrayList<>();
            for (List<String> row : excelRows) {
                excelNames.add(row.get(0));
            }

            excelRows.clear();
            fis.close();
            workbook.close();

            logTextArea.setText("Data imported successfully. Excel Duplicate Count: " + excelDuplicateCount
                    + " and Database Duplicate Count: " + dbDuplicateCount);

        } catch (Exception ex) {
            logTextArea.setText("Error occurred: " + ex.getMessage());
        }
    }

    private String getCellValueAsString(Cell cell) {
        String cellValue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // Convert date cell to a formatted string using a custom date format
                        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                        cellValue = dateFormat.format(cell.getDateCellValue());
                    } else {
                        // Convert numeric cell to a string
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                default:
                    break;
            }
        }
        return cellValue;
    }

    private void storeNonDuplicateRowsInDatabase(List<List<String>> uniqueRows) throws SQLException {
        String url = "jdbc:mysql://localhost:3306/excel_test";
        String username = "root";
        String password = "lynn471997";

        Connection connection = null;
        PreparedStatement preparedStatement = null;

        try {
            connection = DriverManager.getConnection(url, username, password);
            String query = "INSERT INTO employee (employee_name, position, department, salary, joined_date) VALUES (?, ?, ?, ?, ?)";
            preparedStatement = connection.prepareStatement(query);
            for (List<String> value : uniqueRows) {
                preparedStatement.setString(1, value.get(0));
                preparedStatement.setString(2, value.get(1));
                preparedStatement.setString(3, value.get(2));
                preparedStatement.setDouble(4, Double.parseDouble(value.get(3)));
                preparedStatement.setString(5, value.get(4));

                preparedStatement.executeUpdate();
            }
        } catch (SQLException e) {
            System.out.println("Error occurred while inserting data: " + e.getMessage());
            e.printStackTrace();
            throw new RuntimeException(e);
        } finally {
            if (preparedStatement != null) {
                preparedStatement.close();
            }
            if (connection != null) {
                connection.close();
            }
        }
    }


    private List<String> fetchDataFromDatabase() {
        String url = "jdbc:mysql://localhost:3306/excel_test";
        String username = "root";
        String password = "lynn471997";

        List<String> dbRows = new ArrayList<>();

        try (Connection connection = DriverManager.getConnection(url, username, password)) {
            String query = "SELECT employee_name FROM employee";
            try (PreparedStatement preparedStatement = connection.prepareStatement(query);
                 ResultSet resultSet = preparedStatement.executeQuery()) {
                while (resultSet.next()) {
                     String dbRow =resultSet.getString("employee_name");
                      dbRows.add(dbRow);
                }
            }
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
        return dbRows;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ExcelFileReader excelFileReader = new ExcelFileReader();
            excelFileReader.setVisible(true);
        });
    }
}