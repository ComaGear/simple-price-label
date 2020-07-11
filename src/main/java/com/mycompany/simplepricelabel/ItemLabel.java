/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.simplepricelabel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 *
 * @author user
 */
public class ItemLabel implements Comparable<ItemLabel> {
    
    private final String title;
    private final String subhead;
    private final float price;
    private final int sizeType;
    
    public float getTitleSizeInPoint(){
        switch (this.sizeType){
            case ItemLabelStyle.SMALL_SIZE:
                return ItemLabelStyle.ItemStyle.SMALL_TITLE_SIZE;
            case ItemLabelStyle.MEDIUM_SIZE:
                return ItemLabelStyle.ItemStyle.MEDIUM_TITLE_SIZE;
            case ItemLabelStyle.GIANT_SIZE:
                return ItemLabelStyle.ItemStyle.GIANT_TITLE_SIZE;
            default:
                return ItemLabelStyle.ItemStyle.SMALL_TITLE_SIZE;
        }
    }
    public float getSubheadSizeInPoint(){
        switch (this.sizeType){
            case ItemLabelStyle.SMALL_SIZE:
                return ItemLabelStyle.ItemStyle.SMALL_TITLE_SIZE;
            case ItemLabelStyle.MEDIUM_SIZE:
                return ItemLabelStyle.ItemStyle.MEDIUM_TITLE_SIZE;
            case ItemLabelStyle.GIANT_SIZE:
                return ItemLabelStyle.ItemStyle.GIANT_TITLE_SIZE;
            default:
                return ItemLabelStyle.ItemStyle.SMALL_TITLE_SIZE;
        }
    }
    public float getPriceSizeInPoint(){
        switch (this.sizeType){
            case ItemLabelStyle.SMALL_SIZE:
                return ItemLabelStyle.ItemStyle.SMALL_TITLE_SIZE;
            case ItemLabelStyle.MEDIUM_SIZE:
                return ItemLabelStyle.ItemStyle.MEDIUM_TITLE_SIZE;
            case ItemLabelStyle.GIANT_SIZE:
                return ItemLabelStyle.ItemStyle.GIANT_TITLE_SIZE;
            default:
                return ItemLabelStyle.ItemStyle.SMALL_TITLE_SIZE;
        }
    }

    public String getTitle() {
        return title;
    }

    public String getSubhead() {
        return subhead;
    }

    public float getPrice() {
        return price;
    }

    public int getSizeType() {
        return sizeType;
    }
    
    public ItemLabel(String title, String subhead, float price){
        this.title = title;
        this.subhead = subhead;
        this.price = price;
        this.sizeType = ItemLabelStyle.DEFAULT_SIZE;
    }
    
    public ItemLabel(String title, String subhead, float price, int size){
        this.title = title;
        this.subhead = subhead;
        this.price = price;
        this.sizeType = size;
    }

    @Override
    public String toString() {
        return "Title: " + title + "  Subhead: " + subhead + "  price: " + String.valueOf(price) + "  SizeType:" + sizeType;
    }

    @Override
    public int compareTo(ItemLabel o) {
        if(this.sizeType != o.getSizeType()){
            return this.sizeType > o.getSizeType() ? 1 : -1;
        }
        return 0;
    }
    
    
}
