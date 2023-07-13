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
import java.util.Iterator;


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

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new ExcelFileReader().setVisible(true));
    }

    private void importData(File selectedFile) {
        try {
            FileInputStream fis = new FileInputStream(selectedFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            ArrayList<String> excelDataList = new ArrayList<>();
            Iterator<Row> rowIterator = sheet.rowIterator();

            if (rowIterator.hasNext()){
                rowIterator.next();
            }

            while (rowIterator.hasNext()){
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                if (cellIterator.hasNext()){
                    cellIterator.next();
                }

                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    String cellValue = getCellValueAsString(cell);
                    excelDataList.add(cellValue);
                }
            }

            ArrayList<String> databaseDataList = fetchDataFromDatabase();
            ArrayList<String> differentDataList = new ArrayList<>();
            for (String excelData : excelDataList) {
                if (!databaseDataList.contains(excelData)) {
                    differentDataList.add(excelData);
                }
            }

            ArrayList<String> uniqueDataList = new ArrayList<>(differentDataList);

            differentDataList.clear();
            differentDataList.addAll(new HashSet<>(uniqueDataList));

            int duplicateCount = excelDataList.size() - differentDataList.size();

            storeDifferentDataInDatabase(differentDataList);

            logTextArea.setText("Data imported successfully. Duplicate Count : " + duplicateCount);

        } catch (Exception ex) {
            logTextArea.setText("Error occurred : " + ex.getMessage());
        }
    }

    private String getCellValueAsString(Cell cell) {
        String cellValue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                   cellValue = cell.getStringCellValue();
                    System.out.println("CellValue : " + cellValue);
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                       cellValue = cell.getDateCellValue().toString();
                        System.out.println("Date CellValue : " + cellValue);
                        break;
                    } else {
                       cellValue = Double.toString(cell.getNumericCellValue());
                        System.out.println("Salary CellValue : " + cellValue);
                       break;
                    }
                default:
                    break;
            }
        }
        return cellValue;
    }

    private ArrayList<String> fetchDataFromDatabase() throws SQLException {
        ArrayList<String> dataFromDatabase = new ArrayList<>();

        String url = "jdbc:mysql://localhost:3306/excel_test";
        String username = "root";
        String password = "lynn471997";

        Connection connection = null;
        Statement statement = null;
        ResultSet resultSet = null;

        try {
            connection = DriverManager.getConnection(url, username, password);

            statement = connection.createStatement();

            String query = "SELECT employee_name FROM employee";
            resultSet = statement.executeQuery(query);

            while (resultSet.next()) {
                String value = resultSet.getString(1);
                dataFromDatabase.add(value);
            }
        } finally {
            if (resultSet != null) {
                resultSet.close();
            }
            if (statement != null) {
                statement.close();
            }
            if (connection != null) {
                connection.close();
            }
        }
        System.out.println("List of Database : " + dataFromDatabase.size());
        return dataFromDatabase;
    }

    private void storeDifferentDataInDatabase(ArrayList<String> differentDataList) throws SQLException {

        String url = "jdbc:mysql://localhost:3306/excel_test";
        String username = "root";
        String password = "lynn471997";

        Connection connection = null;
        PreparedStatement preparedStatement = null;

        try {
            connection = DriverManager.getConnection(url, username, password);

            String query = "INSERT INTO employee (employee_name, position, department, salary, joined_date) VALUES (?, ?, ?, ?, ?)";
            preparedStatement = connection.prepareStatement(query);

            for (String value : differentDataList) {
                if (value != null && !value.isEmpty()) {
                    preparedStatement.setString(1, value);
                    preparedStatement.setString(2, value);
                    preparedStatement.setString(3, value);
                    preparedStatement.setDouble(4, Double.parseDouble(value));
                    preparedStatement.setDate(5, Date.valueOf(value));
                    preparedStatement.executeUpdate();
                }
            }
        } finally {
            if (preparedStatement != null) {
                preparedStatement.close();
            }
            if (connection != null) {
                connection.close();
            }
        }
    }
}


