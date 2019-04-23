package com.archive.ssm.common.util.EventModel;

import com.archive.ssm.archive.domain.Student;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel2003 implements HSSFListener {

    private SSTRecord sstRecord;

    @Override
    /**
     * getLastCol:每行列数
     */
    public void processRecord(Record record) {
        List<Student> studentList=new ArrayList<>();
        switch (record.getSid()) {
            // the BOFRecord can represent either the beginning of a sheet or the workbook
            case BOFRecord.sid:
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == bof.TYPE_WORKBOOK) {
                    System.out.println("Encountered workbook");
                    // assigned to the class level member
                } else if (bof.getType() == bof.TYPE_WORKSHEET) {
                    System.out.println("Encountered sheet reference");
                }
                break;
            case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                System.out.println("New sheet named: " + bsr.getSheetname());
                break;
            case RowRecord.sid:
                RowRecord rowrec = (RowRecord) record;
                System.out.println("Row found, first column at "
                        + rowrec.getFirstCol() + " last column at " + rowrec.getLastCol());
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;
                System.out.println("Cell found with value " + numrec.getValue()
                        + " at row " + numrec.getRow() + " and column " + numrec.getColumn());
                break;
            // SSTRecords store a array of unique strings used in Excel.
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                for (int k = 0; k < sstRecord.getNumUniqueStrings(); k++) {
                    System.out.println("String table value " + k + " = " + sstRecord.getString(k));
                }
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                System.out.println("String cell found with value "
                        + sstRecord.getString(lrec.getSSTIndex()));
                break;
        }
    }

    public static void main(String[] args) throws IOException{
        // create a new file input stream with the input file specified
        // at the command line
        FileInputStream fin = new FileInputStream("C:\\Users\\DELL\\Downloads\\test.xls");
        // create a new org.apache.poi.poifs.filesystem.Filesystem
        POIFSFileSystem poifs = new POIFSFileSystem(fin);
        // get the Workbook (excel part) stream in a InputStream
        InputStream din = poifs.createDocumentInputStream("Workbook");
        // construct out HSSFRequest object
        HSSFRequest req = new HSSFRequest();
        // lazy listen for ALL records with the listener shown above
        req.addListenerForAllRecords(new ReadExcel2003());
        // create our event factory
        HSSFEventFactory factory = new HSSFEventFactory();
        // process our events based on the document input stream
        factory.processEvents(req, din);
        // once all the events are processed close our file input stream
        fin.close();
        // and our document input stream (don't want to leak these!)
        din.close();
        System.out.println("done.");
    }
}
