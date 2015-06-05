package pe.uni.fiis;


import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;

/**
 * Created by Diego on 5/30/2015.
 */
public class VueloHorario {
    public static void main(String[] args){
        Workbook libro = new HSSFWorkbook();
        Sheet hoja = libro.createSheet();
        Row fila0 = hoja.createRow(0);
        Cell celda01 = fila0.createCell(0);
        Cell celda02 = fila0.createCell(3);
        HSSFRichTextString texto0 = new HSSFRichTextString("Horario-Entrada");
        celda01.setCellValue(texto0);
        HSSFRichTextString texto01 = new HSSFRichTextString("Horario-Salida");
        celda02.setCellValue(texto01);
        Row fila1 = hoja.createRow(1);
        Cell celda10 = fila1.createCell(0);
        Cell celda11 = fila1.createCell(1);
        Cell celda13 = fila1.createCell(3);
        Cell celda14 = fila1.createCell(4);
        HSSFRichTextString texto10 = new HSSFRichTextString("Fecha");
        celda10.setCellValue(texto10);
        HSSFRichTextString texto11 = new HSSFRichTextString("Hora");
        celda11.setCellValue(texto11);
        HSSFRichTextString texto13 = new HSSFRichTextString("Fecha");
        celda13.setCellValue(texto13);
        HSSFRichTextString texto14 = new HSSFRichTextString("Hora");
        celda14.setCellValue(texto14);
        Row fila2 = hoja.createRow(2);
        Cell celda20 = fila2.createCell(0);
        Cell celda21 = fila2.createCell(1);
        Cell celda23 = fila2.createCell(3);
        Cell celda24 = fila2.createCell(4);
        HSSFRichTextString texto20 = new HSSFRichTextString("15-05-15");
        celda20.setCellValue(texto20);
        HSSFRichTextString texto21 = new HSSFRichTextString("10:00 - 11:25");
        celda21.setCellValue(texto21);
        HSSFRichTextString texto23 = new HSSFRichTextString("15-05-15");
        celda23.setCellValue(texto23);
        HSSFRichTextString texto24 = new HSSFRichTextString("12:00 - 13:20");
        celda24.setCellValue(texto24);
        Row fila3 = hoja.createRow(3);
        Cell celda30 = fila3.createCell(0);
        Cell celda31 = fila3.createCell(1);
        Cell celda33 = fila3.createCell(3);
        Cell celda34 = fila3.createCell(4);
        HSSFRichTextString texto30 = new HSSFRichTextString("17-05-15");
        celda30.setCellValue(texto30);
        HSSFRichTextString texto31 = new HSSFRichTextString("8:00 - 9:25");
        celda31.setCellValue(texto31);
        HSSFRichTextString texto33 = new HSSFRichTextString("17-05-15");
        celda33.setCellValue(texto33);
        HSSFRichTextString texto34 = new HSSFRichTextString("17:00 - 18:20");
        celda34.setCellValue(texto34);
        Row fila4 = hoja.createRow(4);
        Cell celda40 = fila4.createCell(0);
        Cell celda41 = fila4.createCell(1);
        Cell celda43 = fila4.createCell(3);
        Cell celda44 = fila4.createCell(4);
        HSSFRichTextString texto40 = new HSSFRichTextString("19-05-15");
        celda40.setCellValue(texto40);
        HSSFRichTextString texto41 = new HSSFRichTextString("12:00 - 13:25");
        celda41.setCellValue(texto41);
        HSSFRichTextString texto43 = new HSSFRichTextString("19-05-15");
        celda43.setCellValue(texto43);
        HSSFRichTextString texto44 = new HSSFRichTextString("15:00 - 16:20");
        celda44.setCellValue(texto44);

        try {
            FileOutputStream elFichero = new FileOutputStream("miexcel.xls");
            libro.write(elFichero);
            elFichero.close();
            System.out.println("miexcel.xls written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
