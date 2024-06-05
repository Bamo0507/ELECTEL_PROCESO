package ExcelEditor;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class ExcelProcessor {
    private static final String EXPORT_FOLDER = "ArchivosModificados";
    private static final int EXPECTED_ROWS = 2977;

    private static final Map<String, String[]> invalidFilesMap = new HashMap<>();

    public static void main(String[] args) {
        // Ruta del directorio que contiene los archivos a modificar
        File folder = new File("ArchivosModificar");

        // Verificar si el directorio existe
        if (!folder.exists()) {
            System.out.println("El directorio especificado no existe.");
            return;
        }

        // Crear el directorio de exportación si no existe
        File exportFolder = new File(EXPORT_FOLDER);
        if (!exportFolder.exists()) {
            exportFolder.mkdirs();
        }

        // Filtrar y obtener solo archivos .xls en el directorio
        File[] listOfFiles = folder.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.toLowerCase().endsWith(".xls"); // Filtra archivos que terminan en .xls
            }
        });

        // Verificar si se encontraron archivos .xls
        if (listOfFiles != null && listOfFiles.length > 0) {
            System.out.println("-----------------------------------------------");
            System.out.println("Archivos encontrados en el directorio:");
            System.out.println("-----------------------------------------------");

            for (File file : listOfFiles) {
                if (file.isFile()) {
                    System.out.println("----------------------");
                    System.out.println("Procesando archivo: " + file.getName());
                    processExcelFile(file); // Procesa cada archivo .xls encontrado
                }
            }

            // Imprimir archivos y hojas con problemas
            printInvalidFilesInfo();
        } else {
            System.out.println("No se encontraron archivos .xls en el directorio.");
        }
    }

    // Método para procesar cada archivo Excel
    private static void processExcelFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new HSSFWorkbook(fis)) { // Abre el archivo Excel

            Sheet sheetToUse = null;
            int maxRowCount = 0;

            // Encontrar la hoja con más filas
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                int rowCount = sheet.getPhysicalNumberOfRows();
                if (rowCount > maxRowCount) {
                    maxRowCount = rowCount;
                    sheetToUse = sheet; // Selecciona la hoja con más filas
                }
            }

            if (sheetToUse != null) {
                System.out.println("Hoja seleccionada: " + sheetToUse.getSheetName());
                System.out.println("----------------------");
                String missingTime = formatSheet(sheetToUse); // Formatea la hoja seleccionada y valida
                exportSheetAsTxt(sheetToUse, file); // Exporta la hoja a un archivo .txt

                if (missingTime != null) {
                    invalidFilesMap.put(file.getName(), new String[]{sheetToUse.getSheetName(), missingTime});
                }
            }
        } catch (IOException e) {
            e.printStackTrace(); // Maneja excepciones de I/O
        }
    }

    // Método para formatear la hoja de cálculo y validar filas
    private static String formatSheet(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle dateCellStyle = workbook.createCellStyle();
        CellStyle numberCellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();

        // Establece el formato de fecha
        dateCellStyle.setDataFormat(dataFormat.getFormat("dd-MM-yyyy HH:mm:ss"));
        // Establece el formato de número con dos decimales
        numberCellStyle.setDataFormat(dataFormat.getFormat("0.00"));

        // Validación y formateo de las filas
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return null;

        int numColumns = headerRow.getPhysicalNumberOfCells();

        if (sheet.getPhysicalNumberOfRows() != EXPECTED_ROWS) {
            return findMissingTime(sheet);
        }

        for (int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            for (int colIndex = 0; colIndex < numColumns; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) continue;

                if (colIndex == 0) {
                    cell.setCellStyle(dateCellStyle); // Aplica estilo de fecha a la primera columna
                } else {
                    if (cell.getCellType() == CellType.NUMERIC) {
                        cell.setCellValue(Math.round(cell.getNumericCellValue() * 100.0) / 100.0); // Redondea a dos decimales
                    }
                    cell.setCellStyle(numberCellStyle); // Aplica estilo de número a las demás columnas
                }
            }
        }
        return null;
    }

    // Método para encontrar la hora faltante
    private static String findMissingTime(Sheet sheet) {
        Set<String> existingTimes = new HashSet<>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");

        for (int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cell = row.getCell(0);
            if (cell != null && cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                existingTimes.add(dateFormat.format(cell.getDateCellValue()));
            }
        }

        Calendar calendar = Calendar.getInstance();
        String missingTime = null;

        // Suponemos que todas las filas pertenecen al mismo día y solo analizamos ese día
        if (!existingTimes.isEmpty()) {
            String firstDateStr = existingTimes.iterator().next();
            Date firstDate = null;
            try {
                firstDate = dateFormat.parse(firstDateStr);
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (firstDate != null) {
                calendar.setTime(firstDate);
                calendar.set(Calendar.HOUR_OF_DAY, 0);
                calendar.set(Calendar.MINUTE, 0);
                calendar.set(Calendar.SECOND, 0);
                calendar.set(Calendar.MILLISECOND, 0);

                for (int i = 0; i < EXPECTED_ROWS - 1; i++) {
                    String time = dateFormat.format(calendar.getTime());
                    if (!existingTimes.contains(time)) {
                        missingTime = time;
                        return missingTime;
                    }
                    calendar.add(Calendar.MINUTE, 15);
                }
            }
        }

        return missingTime;
    }

    // Método para exportar la hoja a un archivo .txt en el directorio de exportación
    private static void exportSheetAsTxt(Sheet sheet, File originalFile) {
        File txtFile = new File(EXPORT_FOLDER, originalFile.getName().replace(".xls", ".txt"));

        try (BufferedWriter bw = new BufferedWriter(new FileWriter(txtFile))) {
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return;

            // Encontrar la columna "timestamp"
            int timestampColIndex = -1;
            for (int colIndex = 0; colIndex < headerRow.getPhysicalNumberOfCells(); colIndex++) {
                Cell cell = headerRow.getCell(colIndex);
                if (cell != null && "timestamp".equalsIgnoreCase(cell.getStringCellValue().trim())) {
                    timestampColIndex = colIndex;
                    break;
                }
            }

            for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

                StringBuilder line = new StringBuilder();
                for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                    if (colIndex == timestampColIndex) {
                        continue; // Omite la columna "timestamp"
                    }
                    Cell cell = row.getCell(colIndex);
                    if (colIndex > 0 && (colIndex != timestampColIndex + 1)) {
                        line.append("\t"); // Añade tabulación entre columnas
                    }
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                line.append(cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    line.append(dateFormat.format(cell.getDateCellValue()));
                                } else {
                                    line.append(String.format("%.2f", cell.getNumericCellValue()));
                                }
                                break;
                            case BOOLEAN:
                                line.append(cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                try {
                                    line.append(cell.getStringCellValue());
                                } catch (IllegalStateException e) {
                                    line.append(cell.getNumericCellValue());
                                }
                                break;
                            default:
                                line.append("");
                        }
                    }
                }
                bw.write(line.toString());
                bw.newLine(); // Nueva línea al final de cada fila
            }
            System.out.println("Archivo exportado como: " + txtFile.getName());
            System.out.println("-----------------------------------------------");
        } catch (IOException e) {
            e.printStackTrace(); // Maneja excepciones de I/O
        }
    }

    // Método para imprimir información sobre archivos y hojas con problemas
    private static void printInvalidFilesInfo() {
        if (invalidFilesMap.isEmpty()) {
            System.out.println("No se encontraron problemas en los archivos procesados.");
        } else {
            System.out.println("Archivos con problemas:");
            System.out.println("-----------------------------------------------");
            for (Map.Entry<String, String[]> entry : invalidFilesMap.entrySet()) {
                String fileName = entry.getKey();
                String[] sheetInfo = entry.getValue();
                System.out.println("Archivo: " + fileName + ", Hoja: " + sheetInfo[0] + ", Hora faltante: " + sheetInfo[1]);
            }
        }
    }
}
