/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.simplepricelabel;

import java.util.List;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 *
 * @author user
 */
public class SheetHandler extends DefaultHandler {

    enum dataType {
        NUMBER, SSTINDEX,
    }

    private final List<ItemLabel> list;
    private String title;
    private String subhead;
    private float price;
    private int size;

    private final SharedStringsTable sst;
    private final StylesTable styles;
    private boolean vIsOpen = false;
    private final StringBuffer value;
    private int thisColumn;

    private dataType nextDataType;
    private int formatIndex;
    private String formatString;
    private final DataFormatter dataFormatter;

    public SheetHandler(SharedStringsTable sst, StylesTable styles, List itemLabelList) {
        this.sst = sst;
        this.styles = styles;
        this.list = itemLabelList;
        this.nextDataType = dataType.NUMBER;
        this.value = new StringBuffer();
        this.dataFormatter = new DataFormatter();
    }

    private int nameToColumn(String name) {
        int column = -1;
        for (int i = 0; i < name.length(); ++i) {
            int c = name.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        String thisStr = null;

        if ("v".equals(qName)) {
            switch (nextDataType) {
                case NUMBER:
                    String n = value.toString();
                    if (this.formatString != null) {
                        thisStr = dataFormatter.formatRawCellContents(Float.parseFloat(n), this.formatIndex, this.formatString);
                    } else {
                        thisStr = n;
                    }
                    break;
                case SSTINDEX:
                    String sstIndex = value.toString();
                    try {
                        int index = Integer.parseInt(sstIndex);
                        XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(index));
                        thisStr = rtss.toString();
                    } catch (NumberFormatException e) {
                    }
                    break;
                default:
                    thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                    break;
            }

            switch (thisColumn) {
                case 0:
                    this.title = thisStr;
                    break;
                case 1:
                    this.subhead = thisStr;
                    break;
                case 2:
                    this.price = Float.parseFloat(thisStr);
                    break;
                case 3:
                    this.size = Integer.parseInt(thisStr);
                    if (Character.isDigit(thisStr.charAt(0))) {
                        this.size = Integer.parseInt(thisStr);
                    }
                    break;
            }
        } else if ("row".equals(qName) && title != null && subhead != null && price != -1f) {
            ItemLabel item = new ItemLabel(title, subhead, price, size);
            list.add(item);
            System.out.println(item.toString());
        }
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {

        if ("v".equals(qName)) {
            vIsOpen = true;
            value.setLength(0);
        } else if ("c".equals(qName)) {

            String r = attributes.getValue("r");
            int firstDigit = 0;
            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }
            thisColumn = nameToColumn(r.substring(0, firstDigit));

            this.nextDataType = dataType.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("s".equals(cellType)) {
                this.nextDataType = dataType.SSTINDEX;
            } else if (cellStyleStr != null) {
                int StyleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = styles.getStyleAt(StyleIndex);
                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();
                if (this.formatString == null) {
                    this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            } else if ("row".equals(qName)) {
                this.title = null;
                this.subhead = null;
                this.price = -1f;
                this.size = 0;
            }
        }
    }
}
