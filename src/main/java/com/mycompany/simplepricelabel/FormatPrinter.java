/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.simplepricelabel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.List;
import java.util.ListIterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormatPrinter {
    
    private final XSSFWorkbook workbook;
    private final ItemLabelStyle labelStyle;
    
    private static final short TITLE_SIZE = 18;
    private static final short SUBHEAD_SIZE = 14;
    private static final short PRICE_SIZE = 24;
    
    private int currentTargetRow = 0;

    public FormatPrinter(XSSFWorkbook workbook, ItemLabelStyle labelStyle){
        this.workbook = workbook;
        this.labelStyle = labelStyle;
    }
    
    // argument Int indentify where sheet to print out, parameter ItemLabel list.
    public XSSFWorkbook process(List<ItemLabel> list){
        XSSFSheet sheet;
        try{
            sheet = workbook.getSheetAt(0);
        }catch(IllegalArgumentException e){
            sheet = workbook.createSheet();
        }
        currentTargetRow = sheet.getLastRowNum();
        currentTargetRow++;
        
        NumberFormat numberF = NumberFormat.getInstance();
        numberF.setMinimumFractionDigits(2);
        numberF.setMaximumFractionDigits(2);
        
        ListIterator iterator = list.listIterator();
        
        boolean isPassRow = false;
        ItemLabel previousItem = null;
        while(isPassRow || iterator.hasNext()) {
            XSSFRow titleRow = sheet.createRow(currentTargetRow++);
            XSSFRow subheadRow = sheet.createRow(currentTargetRow++);
            XSSFRow priceRow = sheet.createRow(currentTargetRow++);
            System.out.println("process row at " + titleRow.getRowNum() + " " + subheadRow.getRowNum() + " " + priceRow.getRowNum() );
            
            boolean columnAvailable = true;
            int targetColumn = 0;
            int numberInRow = 0;
            int lastSize = -1;
            while(isPassRow || iterator.hasNext() && columnAvailable && !isPassRow){
                ItemLabel item;
                if(isPassRow) item = previousItem;
                else item = (ItemLabel) iterator.next();
                if(!(lastSize == -1) && !(lastSize == item.getSizeType())){
                    isPassRow = true;
                    previousItem = item;
                    break;
                }
                
                XSSFCell titleCell = titleRow.createCell(targetColumn);
                XSSFCell subheadCell = subheadRow.createCell(targetColumn);
                XSSFCell priceCell = priceRow.createCell(targetColumn);

                titleCell.setCellValue(item.getTitle());
                labelStyle.TitleStyle(titleCell, item.getSizeType());
                subheadCell.setCellValue(item.getSubhead());
                labelStyle.SubheadStyle(subheadCell, item.getSizeType());
                priceCell.setCellValue("RM " + numberF.format(item.getPrice()));
                labelStyle.PriceStyle(priceCell, item.getSizeType());
                
                System.out.println("cell item at Column " + titleCell.getColumnIndex() + ", value is " + titleCell.getStringCellValue());
                
                numberInRow++;
                lastSize = item.getSizeType();
                isPassRow = false;
                if (numberInRow < ItemLabelStyle.coverMaximum(item.getSizeType())){
                    targetColumn += ItemLabelStyle.coverColumn(item.getSizeType());
                } else {
                    numberInRow = 0;
                    targetColumn = 0;
                    columnAvailable = false;
                }
            }

            // it is process three row at same time. so next to three row to continue.
        }
        
        return workbook;
    }
    
    
    public FileOutputStream print(FileOutputStream output) throws IOException{
        workbook.write(output);
        return output;
    }
    
//    private void defineRow(XSSFRow row, String value, short size){
//        XSSFCell cell = row.createCell(0);
//        cell.setCellStyle(setupStyle(size));
//        cell.setCellValue(value);
//        
//        how to using this method 
////        int itemList = 1;
////        for(int i = 0; i < itemList; i++) {
////            defineRow(sheet.createRow(currentItemRow), "Mini Chipsmore", TITLE_SIZE);
////            currentItemRow++;
////            defineRow(sheet.createRow(currentItemRow), "80g", SUBHEAD_SIZE);
////            currentItemRow++;
////            defineRow(sheet.createRow(currentItemRow), "RM 2.20", PRICE_SIZE);
////            currentItemRow++;
////        }
//    }
    
    private XSSFCellStyle setupStyle(short size){
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        
        font.setFontHeightInPoints(size);
        if (size == PRICE_SIZE) {
            font.setBold(true);
        }
        cellStyle.setFont(font);
        
        return cellStyle;
    }
}
