/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.simplepricelabel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author user
 */
public class ItemLabelStyle {
    
    private final ItemStyle smallItemStyle;
    private final ItemStyle mediumItemStyle;
    private final ItemStyle giantItemStyle;
    
    public static final int SMALL_SIZE = 0;
    public static final int MEDIUM_SIZE = 1;
    public static final int GIANT_SIZE = 2;
    public static final int DEFAULT_SIZE = SMALL_SIZE;
    
    // this is indicate small size that Text is only available in 4 column.
    public static final int SMALL_SIZE_CELL_OVER = 3;
    // this is indicate small size that Text is only available in 5 column.
    public static final int MEDIUM_SIZE_CELL_OVER = 5;
    // this is indicate small size that Text is only available in 9 column.
    public static final int GIANT_SIZE_CELL_OVER = 9;
    // the maximum print out in vertically page.
    public static final int SMALL_SIZE_CELL_MAX = 3;
    // the maximum print out in vertically page.
    public static final int MEDIUM_SIZE_CELL_MAX = 2;
    // the maximum print out in vertically page.
    public static final int GIANT_SIZE_CELL_MAX = 1;
    
    public static int coverColumn(int sizeType){
        switch(sizeType){
            case SMALL_SIZE:
                return SMALL_SIZE_CELL_OVER;
            case MEDIUM_SIZE:
                return MEDIUM_SIZE_CELL_OVER;
            case GIANT_SIZE:
                return GIANT_SIZE_CELL_OVER;
            default :
                return SMALL_SIZE_CELL_OVER;
        }
    }
    
    public static int coverMaximum(int sizeType){
        switch(sizeType){
            case SMALL_SIZE:
                return SMALL_SIZE_CELL_MAX;
            case MEDIUM_SIZE:
                return MEDIUM_SIZE_CELL_MAX;
            case GIANT_SIZE:
                return GIANT_SIZE_CELL_MAX;
            default :
                return SMALL_SIZE_CELL_MAX;
        }
    }
    
    public Cell TitleStyle(Cell cell, int sizeType){
        switch(sizeType){
            case SMALL_SIZE:
                cell.setCellStyle(smallItemStyle.getTitleStyle());
                break;
            case MEDIUM_SIZE:
                cell.setCellStyle(mediumItemStyle.getTitleStyle());
                break;
            case GIANT_SIZE:
                cell.setCellStyle(giantItemStyle.getTitleStyle());
                break;
            default:
                cell.setCellStyle(smallItemStyle.getTitleStyle());
        }
        
        return cell;
    }
    
    public Cell SubheadStyle(Cell cell, int sizeType){
        switch(sizeType){
            case SMALL_SIZE:
                cell.setCellStyle(smallItemStyle.getSubheadStyle());
                break;
            case MEDIUM_SIZE:
                cell.setCellStyle(mediumItemStyle.getSubheadStyle());
                break;
            case GIANT_SIZE:
                cell.setCellStyle(giantItemStyle.getSubheadStyle());
                break;
            default:
                cell.setCellStyle(smallItemStyle.getSubheadStyle());
        }
        
        return cell;
    }
    
    public Cell PriceStyle(Cell cell, int sizeType){
        switch(sizeType){
            case SMALL_SIZE:
                cell.setCellStyle(smallItemStyle.getPriceStyle());
                break;
            case MEDIUM_SIZE:
                cell.setCellStyle(mediumItemStyle.getPriceStyle());
                break;
            case GIANT_SIZE:
                cell.setCellStyle(giantItemStyle.getPriceStyle());
                break;
            default:
                cell.setCellStyle(smallItemStyle.getPriceStyle());
        }
        
        return cell;
    }

    public ItemLabelStyle(Workbook workbook) {
        smallItemStyle = new ItemStyle();
        for(int f = 0; f < 3; f++){
            smallItemStyle.setup(workbook.createCellStyle(), workbook.createFont(), f, SMALL_SIZE);
        }
        
        mediumItemStyle = new ItemStyle();
        for(int f = 0; f < 3; f++){
            mediumItemStyle.setup(workbook.createCellStyle(), workbook.createFont(), f, MEDIUM_SIZE);
        }
        
        giantItemStyle = new ItemStyle();
        for(int f = 0; f < 3; f++){
            giantItemStyle.setup(workbook.createCellStyle(), workbook.createFont(), f, GIANT_SIZE);
        }
    }
    
    public class ItemStyle{
        
        public static final short SMALL_TITLE_SIZE = 18;
        public static final short SMALL_SUBHEAD_SIZE = 14;
        public static final short SMALL_PRICE_SIZE = 24;

        public static final short MEDIUM_TITLE_SIZE = 24;
        public static final short MEDIUM_SUBHEAD_SIZE = 18;
        public static final short MEDIUM_PRICE_SIZE = 32;

        public static final short GIANT_TITLE_SIZE = 36;
        public static final short GIANT_SUBHEAD_SIZE = 24;
        public static final short GIANT_PRICE_SIZE = 48;
        
        private CellStyle titleStyle;
        private CellStyle subheadStyle;
        private CellStyle priceStyle;

        public CellStyle getTitleStyle() {
            return titleStyle;
        }

        public CellStyle getSubheadStyle() {
            return subheadStyle;
        }

        public CellStyle getPriceStyle() {
            return priceStyle;
        }
        
        
        public void setup(CellStyle style, Font font ,int index, int size){
            switch (size){
                case SMALL_SIZE:
                    switch (index){
                        case 0:
                            titleStyle = style;
                            font.setFontHeightInPoints(SMALL_TITLE_SIZE);
                            titleStyle.setFont(font);
                            break;
                        case 1:
                            subheadStyle = style;
                            font.setFontHeightInPoints(SMALL_SUBHEAD_SIZE);
                            subheadStyle.setFont(font);
                            break;
                        case 2:
                            priceStyle = style;
                            font.setFontHeightInPoints(SMALL_PRICE_SIZE);
                            font.setBold(true);
                            priceStyle.setFont(font);
                            break;
                    }
                    break;
                case MEDIUM_SIZE:
                    switch (index){
                        case 0:
                            titleStyle = style;
                            font.setFontHeightInPoints(MEDIUM_TITLE_SIZE);
                            titleStyle.setFont(font);
                            break;
                        case 1:
                            subheadStyle = style;
                            font.setFontHeightInPoints(MEDIUM_SUBHEAD_SIZE);
                            subheadStyle.setFont(font);
                            break;
                        case 2:
                            priceStyle = style;
                            font.setFontHeightInPoints(MEDIUM_PRICE_SIZE);
                            font.setBold(true);
                            priceStyle.setFont(font);
                            break;
                    }
                    break;
                case GIANT_SIZE:
                    switch (index){
                        case 0:
                            titleStyle = style;
                            font.setFontHeightInPoints(GIANT_TITLE_SIZE);
                            titleStyle.setFont(font);
                            break;
                        case 1:
                            subheadStyle = style;
                            font.setFontHeightInPoints(GIANT_SUBHEAD_SIZE);
                            subheadStyle.setFont(font);
                            break;
                        case 2:
                            priceStyle = style;
                            font.setFontHeightInPoints(GIANT_PRICE_SIZE);
                            font.setBold(true);
                            priceStyle.setFont(font);
                            break;
                    }
                    break;
            }
        }

        public ItemStyle() {
        }
        
    }
}
