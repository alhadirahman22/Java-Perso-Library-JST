package PersoJSTLib;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author alhadi.rahman
 */
public class PersoImplement implements PersoInterface {
    private String path;
    private String filename;
    private int split_PerFile;
    private String delimiter;
    private List<String> Data;
    private String[] arrHeaderCol;   
    
    public PersoImplement(String path,String filename,String delimiter,int split, List<String> Data)
    {
       this.path = path;  
       this.filename = filename;
       this.delimiter = delimiter;
       this.split_PerFile = split;
       this.Data = Data;
    }

    @Override
    public void exportTXT() {
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
        // create directory per filename 
        boolean CreatePathFolder = (new File(path + "\\"+"PERSO"+"\\"+filename+"_txtFiles")).mkdirs();
        
        int totalData = Data.size();
        int Splitbagi = split_PerFile;
        int split = totalData/Splitbagi;
        int split2 = totalData%Splitbagi;
        if (split2 > 0) {
            split++;
        }
        int startTake_row = 0;
        for (int j = 1; j < split + 1; j++) {
                if ((split2 > 0) && (j == split) ) {
                    Splitbagi = split2;    
                }
            FileWriter outFile;
            File dstFile = new File(path + "\\"+"PERSO"+"\\"+filename+"_txtFiles"+"\\"+filename+"_"+j+".txt");
            try {
                outFile = new FileWriter(dstFile.getPath());
                PrintWriter writer = new PrintWriter(outFile);
                for (int i = 0; i < Splitbagi; i++) {
                    writer.println(Data.get(startTake_row));    
                    startTake_row++;
                }
            writer.close();
            } catch (IOException ex) {
                Logger.getLogger(PersoImplement.class.getName()).log(Level.SEVERE, null, ex);
            }
        } // end loop Split
    }
    
    // buat header excel
    public void headerExcel(String[] header)
    {
        this.arrHeaderCol = header;
    }
    
    @Override
    public void exportExcel() {
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
        if (this.arrHeaderCol.length > 0) {
            
            boolean CreatePathFolder = (new File(path + "\\"+"PERSO"+"\\"+filename+"_excelFiles")).mkdirs();
            int totalData = Data.size();
            int Splitbagi = split_PerFile;
            int split = totalData/Splitbagi;
            int split2 = totalData%Splitbagi;
            if (split2 > 0) {
                split++;
            }
            int startTake_row = 0;
            String directoryPerso = path + "\\"+"PERSO"+"\\"+filename+"_excelFiles"+"\\";
            for (int j = 1; j < split + 1; j++) {
                XSSFWorkbook xlsxWorkbook=new XSSFWorkbook();
                XSSFSheet sheetxlsxWorkbook=xlsxWorkbook.createSheet("Result");
                XSSFRow rowData =   sheetxlsxWorkbook.createRow((short)0);
                    for (int col = 0; col < this.arrHeaderCol.length; col++) {
                       rowData.createCell((short) col).setCellValue(this.arrHeaderCol[col]); 
                    }
                
                        if ((split2 > 0) && (j == split) ) {
                            Splitbagi = split2;    
                        }
                            int startRow = 1;
                            for (int i = 0; i < Splitbagi; i++) {
                                rowData =   sheetxlsxWorkbook.createRow((short)startRow);
                                String[] temp;
                                temp = Data.get(startTake_row).split(delimiter);
                                for (int k = 0; k < temp.length; k++) { // col ke kanan
                                  rowData.createCell((short) k).setCellValue(temp[k]); 
                                }
                                startTake_row++;
                                startRow++;
                            }
                FileOutputStream fileOut ;
                try {
                    fileOut = new FileOutputStream(directoryPerso+filename+"_"+j+".xlsx");
                    xlsxWorkbook.write(fileOut);
                    fileOut.close();
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(PersoImplement.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(PersoImplement.class.getName()).log(Level.SEVERE, null, ex);
                }
            } // end loop Split
            
        }
        else
        {
            System.out.println("Header belum terisi, Mohon dicek");
        }
    }
    
}
