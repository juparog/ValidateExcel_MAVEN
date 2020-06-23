package com.bancolombia.main;

import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.bancolombia.classes.ExcelFile;
import com.bancolombia.classes.ValidateExcel;

public class Main {
	
	/**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
    	// varibles para el archivo excel
        String filePath = "C:\\Users\\EQUIPO\\Documents\\Aos\\Bancolombia\\Test\\Validar-Excel\\";
        String fileName = "BASILEA_FORWARD2_31122019.xlsx";
        //String fileName = "data.xlsx";
        // String sheet = "31122019";
        int numSheet = 0;
        
        // Se crea un objeto de la clase que lee el archivo
        ExcelFile excelFile = new ExcelFile();
        // Se lee el archivo
        XSSFWorkbook workBook = excelFile.read(filePath + fileName);
        // Se verifica la correcta lectura del archivo
        if(workBook == null){
            System.err.println("Fallo la lectura del libro (ERROR: 101)");
        } else {
            // la lectura del libro ocurrio sin problemas
            // imprimir el contenido de una hoja
            excelFile.printSheet(workBook, numSheet);
            
            // validaciones { {}, {}, {}, {} }
            String [][] validaciones = {
                {"date"}, // columna 1
                {"numeric"}, // columna 2
                {}, // columna 3
                {"string","noOnlyNumeric","strlength-14"} // columna 4
            };
            
            // crear el objeto que valida de la clase que valida el archivo
            ValidateExcel validateExcel = new ValidateExcel();
            int ignoreRows = 1; // filas para ignorar
            
            // ejecutando la validacion
            System.out.println("");
            ArrayList<String> errors = validateExcel.validate(workBook, numSheet, ignoreRows, validaciones);
            
            // imprimiendo el resultado de la validacion
            System.out.println("\n");
            System.out.println("#!#!#! Resultados de la validacion #!#!#!");
            System.out.println("");
            System.out.println("* numero de errores: " + errors.size());
            System.out.println("* impresion del log: ");
            int contError = 0;
            for (String error : errors) {
                System.out.println("    Error #" + contError++ + ": " + error);
            }
            
            // cargar en la DB
            // out(fila) || error bull copy
        }
    }

}
