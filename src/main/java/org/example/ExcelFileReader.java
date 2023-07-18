package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.sql.*;
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
            Set<List<String>> uniqueRows = new HashSet<>(); // Set to store unique rows

            int excelDuplicateCount = 0;
            int dbDuplicateCount = 0;

            // Iterate over rows
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                List<String> rowData = new ArrayList<>(); // ArrayList to store the cells of each row

                // Iterate over cells in the row
                for (Cell cell : row) {
                    // Get the cell value as a String
                    String cellValue = getCellValueAsString(cell);
                    rowData.add(cellValue);
                    System.out.println("Row Data: " + rowData);
                }

                // Check if the row is duplicate based on relevant columns
                boolean isDuplicate = false;
                for (List<String> existingRow : uniqueRows) {
                    // Compare relevant columns (e.g., excluding the ID column)
                    if (existingRow.subList(1, existingRow.size()).equals(rowData.subList(1, rowData.size()))) {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate) {
                    uniqueRows.add(rowData);
                    excelRows.add(rowData);
                } else {
                    excelDuplicateCount++;
                }
            }
            System.out.println("Unique Rows: " + uniqueRows);
            System.out.println("Excel Rows: " + excelRows);

            List<String> rowsFromDb = fetchDataFromDatabase();
            List<List<String>> nonDuplicateRows = new ArrayList<>();

            if (rowsFromDb.isEmpty()){
                storeNonDuplicateRowsInDatabase(excelRows);
            }

            for (List<String> row : excelRows) {
                if (!rowsFromDb.contains(row.get(1))) {
                    nonDuplicateRows.add(row);
                } else {
                    dbDuplicateCount++;
                }
            }

            storeNonDuplicateRowsInDatabase(nonDuplicateRows);

            excelRows.clear();
            uniqueRows.clear();
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
                        cellValue = cell.getDateCellValue().toString();
                    } else {
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
            connection.setAutoCommit(true);
            String query = "INSERT INTO employee (employee_name, position, department, salary, joined_date) VALUES (?, ?, ?, ?, ?)";
            preparedStatement = connection.prepareStatement(query);
            for (List<String> value : uniqueRows) {
                    preparedStatement.setString(1, value.get(1));
                    preparedStatement.setString(2, value.get(2));
                    preparedStatement.setString(3, value.get(3));
                    preparedStatement.setDouble(4, Double.parseDouble(value.get(4)));

                    String dateString = value.get(5);
                    Date date;
                    try{
                        date = Date.valueOf(dateString);
                    }catch (IllegalArgumentException ex) {
                        logTextArea.setText("Invalid date format " + dateString );
                        return;
                    }
                    preparedStatement.setDate(5, date);
                    preparedStatement.executeUpdate();
            }
            connection.commit();
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