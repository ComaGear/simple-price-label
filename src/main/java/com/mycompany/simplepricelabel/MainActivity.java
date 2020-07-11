/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.simplepricelabel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

public class MainActivity {
    
    private static final String outputFile = "D:\\Hard Working\\ProjectData\\PriceLabel.xlsx";
    private static final String inputFile = "D:\\Hard Working\\ProjectData\\data.xlsx";

    public static void main(String[] args){
        try {
            String outputFileName = args[0];
            String inputFileName = args[1];
            File outputFile = new File(outputFileName);
            File inputFile = new File(inputFileName);
            XSSFWorkbook outputWorkbook = createOutputFile(outputFile);
            FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
            
            List<ItemLabel> itemLabels = new LinkedList<ItemLabel>();
            itemLabels = parseDataSource(inputFile, itemLabels);
            Collections.sort(itemLabels);
            
            for(ItemLabel item : itemLabels){
                System.out.println(item);
            }

            ItemLabelStyle styleList = new ItemLabelStyle(outputWorkbook);
            FormatPrinter formatPrinter = new FormatPrinter(outputWorkbook, styleList);
            formatPrinter.process(itemLabels);
            formatPrinter.print(fileOutputStream)
                    .close();
        } catch (IOException | OpenXML4JException | SAXException | ParserConfigurationException ex) {
            Logger.getLogger(MainActivity.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static List<ItemLabel> parseDataSource(File input, List<ItemLabel> list) throws InvalidFormatException, OpenXML4JException, IOException, SAXException, ParserConfigurationException{
        OPCPackage opg = OPCPackage.open(input);
        XSSFReader reader = new XSSFReader(opg);
        SharedStringsTable sst = reader.getSharedStringsTable();
        StylesTable styles = reader.getStylesTable();
        XMLReader parser = XMLHelper.newXMLReader();
        ContentHandler contentHandler = new SheetHandler(sst, styles, list);
        parser.setContentHandler(contentHandler);
        InputStream inputStream = reader.getSheetsData().next();
        InputSource source = new InputSource(inputStream);
        parser.parse(source);
        
        return list;
    }
    
    private static XSSFWorkbook createOutputFile(File file) throws IOException {
        XSSFWorkbook workbook;
        FileInputStream fileInputStream;
        if (file.createNewFile()) {
            System.out.println("File is already created.");
            workbook = new XSSFWorkbook();
        } else {
            System.out.println("File is exists.");
            fileInputStream = new FileInputStream(file);
            workbook = (XSSFWorkbook) WorkbookFactory.create(fileInputStream);
        }
        return workbook;
    }
}
