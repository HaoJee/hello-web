package com.archive.ssm.common.util.EventModel;

import org.apache.poi.util.BoundedInputStream;
import org.springframework.ui.Model;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.*;
import java.nio.channels.Channels;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Test {

    private File file=new File("C:\\Users\\DELL\\Downloads\\testData\\档案馆系统1w.xlsx");
    private int countStream=0;
    private int chunkSize=1024*100;
    public static void main(String[] args){

        Test t=new Test();
        try {
            List<InputStream> list=t.getAllInputStreams();
            int i=0;
            for (InputStream in: list) {
                OutputStream outputStream=new FileOutputStream("C:\\Users\\DELL\\Downloads\\testData\\new"+i+".xlsx");

                while(in.read()!=-1){
                    outputStream.write(in.read());
                }
                i++;
            }
            System.out.println("pppp");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    protected InputStream createInputStream() throws Exception{
        RandomAccessFile randomAccessFile = new RandomAccessFile(this.file, "r");
        BoundedInputStream res = new BoundedInputStream(
                Channels.newInputStream(randomAccessFile.getChannel().position(this.countStream * chunkSize)), chunkSize);
        res.setPropagateClose(false) ;
        return res ;
    }

    public ArrayList<InputStream> getAllInputStreams() throws Exception{
        ArrayList<InputStream> allStreams = new ArrayList<InputStream>();
        InputStream stream = this.getNext();
        while (stream != null) {
            allStreams.add(stream);
            stream = this.getNext();
        }
        return allStreams;
    }

    public InputStream getNext() throws Exception{
        InputStream segment = this.createInputStream();
        this.countStream++;
        return segment;
    }

}
