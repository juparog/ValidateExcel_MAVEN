package com.bancolombia.classes;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFile {
	
	public ExcelFile() {}
	
	/**
     * Metodo para leer un archivo excel
     * @param pathFile, ruta del archivo a leer
     * @return workBook || null, libro excel o nulo si algo falla
     */
    public XSSFWorkbook read(String pathFile){
        try {
            // Se lee el archivo
            File excelFile = new File(pathFile);
            // Se procesa el archivo
            InputStream excelStream = new FileInputStream(excelFile);
            // Se crea el libro de trabajo a partir de la lectura del archivo
            XSSFWorkbook workBook = new XSSFWorkbook(excelStream);
            // Se retorna el libro de excel
            return workBook;
        } catch (FileNotFoundException ex) {
            System.err.println("No se encontró el fichero (ERROR: 201): " + ex);
        } catch (IOException ex) {
            System.err.println("Error al procesar el fichero (ERROR: 202): " + ex);
        }
        return null;
    }
    
    /**
     * Metodo que imprime una hoja
     * @param workBook, libro de excel
     * @param numSheet, posicion de la hoja en el libro
     */
    public void printSheet(XSSFWorkbook workBook,int numSheet){
        System.out.println("#!#!#! Imprimiendo hoja numero: " + numSheet + " #!#!#!");
        System.out.println("");
        // se obtiiene la hoja del libro para imprimir
        XSSFSheet sheet = workBook.getSheetAt(numSheet);
        // numero de filas a imprimir
        int totalRows = sheet.getLastRowNum();
        // se llama a la funcion de la clase que imprime las filas
        this.printRows(sheet, totalRows);
    }
    
    /**
     * Metodo para imprimir filas de una hoja
     * @param sheet, hoja del libro
     * @param numRowsPrint, numero de filas a imprimir
     */
    private void printRows(XSSFSheet sheet, int numRowsPrint){
        System.out.println("Numero de filas a imprimir: " + numRowsPrint);
        // ciclo que recorre todas las filas de la hoja
        for (int numRow = 0; numRow <= numRowsPrint; numRow++) {
            // se obtiene la fila segun la posicion del ciclo
            XSSFRow xssfRow = sheet.getRow(numRow);
            if (xssfRow == null){ // entra en el caso de que la fila este nula
                System.out.println("Fila " + (numRow + 1) + " -> [Fila nula ...]");
                // break;
            }else{
                System.out.print("Fila " + (numRow + 1) + " -> ");
                for (int numCol = 0; numCol < xssfRow.getLastCellNum(); numCol++) {
                    String cellValue;
                    cellValue = xssfRow.getCell(numCol) == null?"":
                            (xssfRow.getCell(numCol).getCellType() == CellType.STRING)?xssfRow.getCell(numCol).getStringCellValue():
                            (xssfRow.getCell(numCol).getCellType() == CellType.NUMERIC)?"" + xssfRow.getCell(numCol).getNumericCellValue():
                            (xssfRow.getCell(numCol).getCellType() == CellType.BOOLEAN)?"" + xssfRow.getCell(numCol).getBooleanCellValue():
                            (xssfRow.getCell(numCol).getCellType() == CellType.BLANK)?"$$$BLANK$$$":
                            (xssfRow.getCell(numCol).getCellType() == CellType.FORMULA)?"$$$FORMULA$$$":
                            (xssfRow.getCell(numCol).getCellType() == CellType.ERROR)?"$$$ERROR$$$":
                            (xssfRow.getCell(numCol).getCellType() == CellType._NONE)?"$$$ERROR$$$":"$$$NONE$$$";                       
                    System.out.print("[Columna " + numCol + ": " + cellValue + "] ");
                }
                System.out.println("");
            }
        }
    }
	
}
