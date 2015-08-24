/*
 * Copyright (C) 2015  Ján Trenčanský
 *
 *    This program is free software; you can redistribute it and/or modify
 *    it under the terms of the GNU General Public License as published by
 *    the Free Software Foundation; either version 3 of the License, or
 *    (at your option) any later version.
 *
 *    This program is distributed in the hope that it will be useful,
 *    but WITHOUT ANY WARRANTY; without even the implied warranty of
 *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *    GNU General Public License for more details.
 *
 *    You should have received a copy of the GNU General Public License
 *    along with this program; if not, write to the Free Software Foundation,
 *    Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301  USA
 */

import java.io.*;

import org.apache.commons.cli.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {


    public static void main(String[] args) {
        String input = null;
        String output = null;
        String separator = ",";
        String encodingInput = "utf-8";
        String encodingOutput = null;

        //Create commandline parser
        CommandLineParser parser = new DefaultParser();

        //Create help formatter
        HelpFormatter formatter = new HelpFormatter();

        //Create options
        Options options = new Options();

        options.addOption("i", "input", true, "input csv file");
        options.addOption("o", "output", true, "output csv file");
        options.addOption("s", "separator", true, "specify separator default(,)");
        options.addOption("h", "help", false, "print help");
        options.addOption("ei", "encoding-input", true, "input csv file encoding default(utf-8)");
        options.addOption("eo", "encoding-output", true, "output csv file encoding default(system)");

        try {
            CommandLine line = parser.parse(options, args);

            if(line.hasOption("h")){
                formatter.printHelp("csv2xlsx -i <input.csv> -o <output.xlsx>", options);
                System.exit(0);
            }

            if(line.hasOption("i")){
                input = line.getOptionValue("i");
            } else {
                formatter.printHelp("please specify input:", options);
                System.exit(0);
            }

            if(line.hasOption("o")){
                output = line.getOptionValue("o");
            } else {
                formatter.printHelp("please specify output:", options);
                System.exit(0);
            }

            if(line.hasOption("s")){
                separator = line.getOptionValue("s");
            }

            if(line.hasOption("ei")){
                encodingInput = line.getOptionValue("ei");
            }

            if(line.hasOption("eo")){
                encodingOutput = line.getOptionValue("eo");
            }

            if((input != null) && (output != null)){
                try {
                    FileInputStream inputTest = new FileInputStream(input);
                    InputStreamReader reader = new InputStreamReader(inputTest, encodingInput);
                    FileOutputStream outputTest = new FileOutputStream(input+"_temp");
                    OutputStreamWriter writer;
                    if(encodingOutput != null){
                        writer = new OutputStreamWriter(outputTest, encodingOutput);
                    } else {
                        writer = new OutputStreamWriter(outputTest);
                    }
                    //writer.write("\uFEFF");
                    int read = reader.read();
                    while (read != -1) {
                        writer.write(read);
                        read = reader.read();
                    }
                    reader.close();
                    writer.close();
                    inputTest.close();
                    outputTest.close();
                    csvToXLSX(input+"_temp", output, separator);
                } catch(Exception e){
                    System.out.println(e.toString());
                }
            }

        } catch (ParseException e){
            System.out.println("Parsing failed:" + e.getMessage());
            formatter.printHelp("csv2xlsx -i <input.csv> -o <output.xlsx>", options);
        }
    }

    public static void csvToXLSX(String csvFile, String xlsxFile, String separator) {
        try {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("sheet1");
            String currentLine;
            int RowNum=0;
            BufferedReader br = new BufferedReader(new FileReader(csvFile));
            while ((currentLine = br.readLine()) != null) {
                String str[] = currentLine.split(separator);
                XSSFRow currentRow=sheet.createRow(RowNum);
                RowNum++;
                for(int i=0;i<str.length;i++){
                    currentRow.createCell(i).setCellValue(str[i]);
                }
            }

            FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFile);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Conversion done");
            System.out.println("Output:" + xlsxFile);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

}