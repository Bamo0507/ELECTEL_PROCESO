package ExcelEditor;

//Librerías importadas
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelProcessor {
    private static final String EXPORT_FOLDER = "ProcesadorExcel\\ArchivosModificados"; // Se debe modificar el PATH hacia el fólder donde se quieren los txts
    private static final Map<String, String[]> invalidFilesMap = new HashMap<>();

    public static void main(String[] args) {
        // Ruta del directorio que contiene los archivos a modificar
        File folder = new File("ProcesadorExcel\\ArchivosModificar"); // Se debe modificar el PATH hacia el fólder donde se tengan los libros de Excel

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

        // Filtrar y obtener solo archivos .xls y .xlsx en el directorio
        File[] listOfFiles = folder.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx");
            }
        });

        // Verificar si se encontraron archivos .xls o .xlsx
        if (listOfFiles != null && listOfFiles.length > 0) {
            System.out.println("-----------------------------------------------");
            System.out.println("Archivos encontrados en el directorio:");
            System.out.println("-----------------------------------------------");

            for (File file : listOfFiles) {
                if (file.isFile()) {
                    System.out.println("----------------------");
                    System.out.println("Procesando archivo: " + file.getName());
                    processExcelFile(file); // Procesa cada archivo .xls o .xlsx encontrado
                }
            }

            // Imprimir archivos y hojas con problemas
            printInvalidFilesInfo();

            // Corregir extensiones .txtx a .txt
            fixTxtxExtensions();

        } else {
            System.out.println("No se encontraron archivos .xls o .xlsx en el directorio.");
        }
    }

    // Corregir extensiones .txtx a .txt
    private static void fixTxtxExtensions() {
        File exportFolder = new File(EXPORT_FOLDER);
        File[] txtxFiles = exportFolder.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File dir, String name) {
                return name.toLowerCase().endsWith(".txtx");
            }
        });

        if (txtxFiles != null) {
            for (File txtxFile : txtxFiles) {
                String txtxFileName = txtxFile.getName();
                String txtFileName = txtxFileName.substring(0, txtxFileName.length() - 1); // Eliminar la última 'x'
                File txtFile = new File(exportFolder, txtFileName);
                
            }
        }
    }


    // Método para procesar cada archivo Excel
    private static void processExcelFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = file.getName().toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis)) { // Abre el archivo Excel

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
                int expectedRows = calculateExpectedRows(sheetToUse); // Calcula la cantidad esperada de filas basado en el mes y año
                String missingTime = formatSheet(sheetToUse, expectedRows); // Formatea la hoja seleccionada y valida
                exportSheetAsTxt(sheetToUse, file); // Exporta la hoja a un archivo .txt

                if (missingTime != null) {
                    invalidFilesMap.put(file.getName(), new String[]{sheetToUse.getSheetName(), missingTime});
                }
            }
        } catch (IOException e) {
            System.err.println("Error al procesar el archivo " + file.getName() + ": " + e.getMessage());
        }
    }

    // Método para calcular el número esperado de filas basado en el mes y año
    private static int calculateExpectedRows(Sheet sheet) {
        Calendar calendar = Calendar.getInstance();
        Date date = sheet.getRow(1).getCell(0).getDateCellValue();
        calendar.setTime(date);
        int daysInMonth = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
        return daysInMonth * 24 * 4 + 1; // Número de días * 24 horas * 4 (cada 15 minutos) + 1 fila de encabezado
    }

    // Método para formatear la hoja de cálculo y validar filas
    private static String formatSheet(Sheet sheet, int expectedRows) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle dateCellStyle = workbook.createCellStyle();
        CellStyle numberCellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();

        // Establece el formato de fecha para las dos primeras columnas
        dateCellStyle.setDataFormat(dataFormat.getFormat("dd-MM-yyyy HH:mm:ss"));
        // Establece el formato de número con dos decimales
        numberCellStyle.setDataFormat(dataFormat.getFormat("0.00"));

        // Validación y formateo de las filas
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return null;

        int numColumns = headerRow.getPhysicalNumberOfCells();

        if (sheet.getPhysicalNumberOfRows() != expectedRows) {
            return findMissingTime(sheet, expectedRows);
        }

        for (int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            for (int colIndex = 0; colIndex < numColumns; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) continue;

                if (colIndex == 0 || colIndex == 1) { // Formato de fecha solo para las dos primeras columnas
                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                        cell.setCellStyle(dateCellStyle); // Aplica estilo de fecha a las dos primeras columnas
                    } 
                } else { // Formato de número para el resto de las columnas
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
    private static String findMissingTime(Sheet sheet, int expectedRows) {
        Set<String> existingTimes = new HashSet<String>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");

        // Recorre las filas y almacena las fechas de la primera columna
        for (int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cell = row.getCell(0); // Analiza solo la primera columna
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
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

                for (int i = 0; i < expectedRows - 1; i++) {
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
        String txtFileName = originalFile.getName().replace(".xls", ".txt").replace(".xlsx", ".txt");
        File txtFile = new File(EXPORT_FOLDER, txtFileName);

        try (BufferedWriter bw = new BufferedWriter(new FileWriter(txtFile))) {
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return;

            for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

                StringBuilder line = new StringBuilder();
                for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if (colIndex > 0) {
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
            System.err.println("Error al exportar el archivo " + originalFile.getName() + " como TXT: " + e.getMessage());
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
