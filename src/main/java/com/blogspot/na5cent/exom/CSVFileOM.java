package com.blogspot.na5cent.exom;

import com.blogspot.na5cent.exom.util.EachFieldCallback;
import com.blogspot.na5cent.exom.util.ReflectionUtils;
import com.blogspot.na5cent.exom.util.StringReturnCallBack;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;

import java.lang.reflect.Field;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;


public class CSVFileOM {
        
    private final File csvFile;
    private Class clazz;
    private String seperator = ",";
    
    
    private CSVFileOM(File csvFile) {
        this.csvFile = csvFile;
    }
    
    public static CSVFileOM mapFromCsvFile(File csvFile) {
        return new CSVFileOM(csvFile);
    }
    
    
    
    public CSVFileOM to(Class clazz) {
        this.clazz = clazz;
        return this;
    }
    
    
    public CSVFileOM csvSeperator(String seperator) {
        this.seperator = seperator;
        return this;
    }
    
    
    
    public <T> List<T> map() throws Throwable {
        InputStream inputStream = null;
        List<T> items = new LinkedList<T>();
        BufferedReader br = null;

        try {
            br = new BufferedReader(new FileReader(csvFile));
            String line = "";

            int i = 0;
            Map<String, Integer> nameIndexMap = new HashMap<String, Integer>();
            while ((line = br.readLine()) != null) {

                if (i == 0) {
                    readFileHeader(line, nameIndexMap);
                } else {
                    items.add((T) readFileContent(line, nameIndexMap));
                }
                i++;


            }
           
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }

        return items;
    }
    
    

    private void mapName2Index(String name, String line, Map<String, Integer> cells) {
        int index = findIndexCellByName(name, line);
        if (index != -1) {
            cells.put(name, index);
        }
    }
    
    
    private int findIndexCellByName(String name, String line) {
       
       String[] lineCols = line.split(this.seperator);
       int i = 0;
       for(String column : lineCols) {
           if(column != null && column.equals(name)){
               return i;
           }
           i++;
       }
        return -1;
    }

    private void readFileHeader(final String line, final Map<String, Integer> cells)  throws Throwable {
        
        ReflectionUtils.eachFields(clazz, new EachFieldCallback() {

            @Override
            public void each(Field field, String name) throws Throwable {
                mapName2Index(name, line, cells);
            }
        });
        
    }

    private Object readFileContent(final String line, final Map<String, Integer> cells) throws Throwable {
        final Object instance = clazz.newInstance();
        ReflectionUtils.eachFields(clazz, new EachFieldCallback() {

            @Override
            public void each(Field field, String columnName) throws Throwable {
                String cellValue = getCellValue(columnName, line, cells);
                ReflectionUtils.setValueOnField(instance, field,cellValue);
            }


          
        });
    return instance;
    }
    
    private String getCellValue(String columnName, String line, Map<String, Integer> cells) {
        if (cells.get(columnName) == null) {
            return null;
        }
        int pos = cells.get(columnName);
        String[] lineCols = line.split(this.seperator);
        
        if(lineCols != null && lineCols.length > pos){
            return lineCols[pos];
        }else {
            return null;
        }
        
        
    }
    
    
    public <T> void write(List<T> dataList) throws Throwable {
        BufferedWriter bw = null;

        try {
           
            bw = new BufferedWriter(new FileWriter(csvFile, false));
            writeHeader(bw);
            bw.newLine();
            for (T data : dataList) {
                   writeDataLine(data, bw);
                    bw.newLine();
                    bw.flush();
                }
           
           
        } finally {
           bw.close();
        }

    }

    private void writeHeader(BufferedWriter writer) throws IOException {
        
        List<String> headerList = ReflectionUtils.getFields(clazz);
        StringBuffer headerRow = new StringBuffer();
        String sep = "";
        
            for(String column : headerList){
                headerRow.append(sep).append(column);
                sep = this.seperator;
            }
        writer.write(headerRow.toString());
    }

    private  void  writeDataLine(final Object data, BufferedWriter bw) throws Throwable {
        StringBuffer lineRow = new StringBuffer();
        List<String> dataList = ReflectionUtils.listFieldValue(clazz, new StringReturnCallBack() {
                @Override
                public String each(Field field, String columnName) throws Throwable {
                    return ReflectionUtils.getValueOfField(data, field);
                }
            
            
            });
        String sep = "";
        
            for(String column : dataList){
                lineRow.append(sep).append(column);
                sep = this.seperator;
            }
        
        bw.write(lineRow.toString());
        
    }
}
