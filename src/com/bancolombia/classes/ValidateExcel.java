package com.bancolombia.classes;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author EQUIPO
 */
public class ValidateExcel {

    public ValidateExcel() {
    }
    
    public ArrayList<String> validate(XSSFWorkbook workbook, int numSheet, int ignoreRows, String [][] validaciones){
        System.out.println("#!#!#! Validando archivo, numero de hoja: " + numSheet + " #!#!#!");
        System.out.println("");
        
        ArrayList<String> errors = new ArrayList<>(); // variable para contar los errores
        try {        
            XSSFSheet sheet = workbook.getSheetAt(numSheet);
            int totalRows = sheet.getLastRowNum();

            // ciclo para leer las filas
            for (int numRow = ignoreRows; numRow <= totalRows; numRow++) {
                System.out.println("Fila: " + numRow);
                XSSFRow xssfRow = sheet.getRow(numRow);
                if (xssfRow == null){
                    errors.add("El archivo contiene una fila vacia entre los datos(ERROR: 301), posición: " + (numRow + ignoreRows));
                }else{
                    for (int numCol = 0; numCol < xssfRow.getLastCellNum(); numCol++) {
                        // { {"date"}, {"numeric"}, {}, {"string","strlength-10"}}  [0-9]*
                        try {
                            System.out.print("    Columna: " + numCol + " [");
                            
                            // Obtener la validacio segun la columna
                            String[] validacionesCol = validaciones[numCol];
                            for (String strVal : validacionesCol) {
                                System.out.print(" " + strVal + " , ");
                                // Validando
                                // String cellValue = this.getValueCell(xssfRow, numCol);
                                if(strVal.equals("date")){
                                    if(!this.isDate(xssfRow, numCol)){
                                        errors.add("No cumple la validacion \'date\' (ERROR: 302), posición: " + (numRow + ignoreRows));
                                    }
                                } else if(strVal.equals("numeric")){
                                    if(!this.isNumeric(xssfRow, numCol)){
                                        errors.add("No cumple la validacion \'numeric\' (ERROR: 303), posición: " + (numRow + ignoreRows));
                                    }
                                } else if(strVal.equals("noOnlyNumeric")){
                                    if(!this.noOnlyNumeric(xssfRow, numCol)){
                                        errors.add("No cumple la validacion \'noOnlyNumeric\' (ERROR: 304), posición: " + (numRow + ignoreRows));
                                    }
                                } else if(strVal.matches("^strlength-[0-9]*")){
                                    if(!this.strLength(xssfRow, numCol, strVal)){
                                        errors.add("No cumple la validacion \'strlength-0\' (ERROR: 305), posición: " + (numRow + ignoreRows));
                                    }
                                }
                            }
                            System.out.println(" ]");
                        } catch (ArrayIndexOutOfBoundsException ex) {
                            errors.add("El numero de columnas en el archivo no concuerda con la "
                                    + "definidas en la validacion, posicion del registro: " 
                                    + (numRow + ignoreRows) + ", (ERROR: 303): " + ex.getMessage());
                            System.out.println(" ]");
                            break;
                        }
                    }
                }
            }
        } catch (Exception ex) {
            errors.add("Error en la validación del archivo (ERROR: 304):" + ex.getMessage());
        }
        return errors;
    }
    
    private String getValueCell(XSSFRow xssfRow, int numCol){
        String cellValue = xssfRow.getCell(numCol) == null?"":
                (xssfRow.getCell(numCol).getCellType() == CellType.STRING)?xssfRow.getCell(numCol).getStringCellValue():
                (xssfRow.getCell(numCol).getCellType() == CellType.NUMERIC)?"" + xssfRow.getCell(numCol).getNumericCellValue():
                (xssfRow.getCell(numCol).getCellType() == CellType.BOOLEAN)?"" + xssfRow.getCell(numCol).getBooleanCellValue():
                (xssfRow.getCell(numCol).getCellType() == CellType.BLANK)?"$$$-BLANK-$$$":
                (xssfRow.getCell(numCol).getCellType() == CellType.FORMULA)?"$$$-FORMULA-$$$":
                (xssfRow.getCell(numCol).getCellType() == CellType.ERROR)?"$$$-ERROR-$$$":
                (xssfRow.getCell(numCol).getCellType() == CellType._NONE)?"$$$-NONE-$$$":"$$$-UNDEFINED-$$$";
        return cellValue;
    }
    
    private boolean isDate(XSSFRow xssfRow, int numCol){
        try {
            String date = xssfRow.getCell(numCol).getDateCellValue().toString();
            return true;
        } catch (Exception e) {
            // System.err.println(e.getMessage());
            return false;
        }
    }
    
    private boolean isNumeric(XSSFRow xssfRow, int numCol){
        try {
            double num = xssfRow.getCell(numCol).getNumericCellValue();
            return true;
        } catch (Exception e) {
            // System.err.println(e.getMessage());
            return false;
        }
    }
    
    private boolean isString(XSSFRow xssfRow, int numCol){
        try {
            String str = xssfRow.getCell(numCol).getStringCellValue();
            return true;
        } catch (Exception e) {
            // System.err.println(e.getMessage());
            return false;
        }
    }
    
    private boolean noOnlyNumeric(XSSFRow xssfRow, int numCol){
        try {
            String str = xssfRow.getCell(numCol).getStringCellValue();
            if(!str.matches("[0-9]*")){
                return true;
            }
            return false;
        } catch (Exception e) {
            // System.err.println(e.getMessage());
            return false;
        }
    }
    
    private boolean strLength(XSSFRow xssfRow, int numCol, String strVal){
        try {
            String[] arr = strVal.split("-");
            String str = xssfRow.getCell(numCol).getStringCellValue();
            if(str.length() == Integer.parseInt(arr[1])){
                return true;
            }
            return false;
        } catch (Exception e) {
            // System.err.println(e.getMessage());
            return false;
        }
    }
}

