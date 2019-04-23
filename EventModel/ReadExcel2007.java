package com.archive.ssm.common.util.EventModel;

import com.archive.ssm.common.util.ExcelPC.TransDataUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class ReadExcel2007 extends DefaultHandler{

    public void processOneSheet(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst);

        // To look up the Sheet Name / Sheet Order / rID,
        //  you need to process the core Workbook stream.
        // Normally it's of the form rId# or rSheet#
        InputStream sheet1 = r.getSheet("rId1");
        InputSource sheetSource = new InputSource(sheet1);
        parser.parse(sheetSource);
        System.out.println("11111111");
        sheet1.close();
    }

    public void processAllSheets(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst);

        Iterator<InputStream> sheets = r.getSheetsData();
        while(sheets.hasNext()) {
            System.out.println("\nProcessing new sheet:");
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
            System.out.println("");
        }
        List<List<Object>> studnets=SheetHandler.studentList;
        for (List<Object> l:studnets){
            for (Object o:l) {
                System.out.print(String.valueOf(o)+" ");
            }
            System.out.println();
        }
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser =
                XMLReaderFactory.createXMLReader(
                        "org.apache.xerces.parsers.SAXParser"
                );
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    private static class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private List<Object> rowValue=new ArrayList<>();
        private static List<List<Object>> studentList=new ArrayList<>();
        private String lastContents;
        private boolean nextIsString;

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            // c => cell
            if(name.equals("c")) {
                // Print the cell reference
//                System.out.print(attributes.getValue("r") + " - ");
                // Figure out if the value is an index in the SST
                String cellType = attributes.getValue("t");
                if(cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }
            // Clear contents cache
            lastContents = "";
        }

        public void endElement(String uri, String localName, String name)
                throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                nextIsString = false;
                rowValue.add(lastContents);
                if(idx!=0 && (idx+1)%8==0){
//                    TransDataUtil
                    studentList.add(new ArrayList<>(rowValue));
                    rowValue.clear();
                }
//                System.out.println(idx+" "+sst.getEntryAt(idx)+" "+new XSSFRichTextString(sst.getEntryAt(idx)));
            }

            // v => contents of a cell
            // Output after we've seen the string contents
//            if(name.equals("v")) {
//                System.out.println("cell value："+lastContents);
//            }
        }

        public void characters(char[] ch, int start, int length)
                throws SAXException {
//            System.out.println("characters: "+new String(ch, start, length)+"  "+start+" "+length);
            lastContents += new String(ch, start, length);
        }
    }

    public static void main(String[] args) throws Exception {
        ReadExcel2007 example = new ReadExcel2007();
//        example.processOneSheet("C:\\Users\\DELL\\Downloads\\test.xlsx");
        example.processAllSheets("C:\\Users\\DELL\\Downloads\\test.xlsx");
    }
}
