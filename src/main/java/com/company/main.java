package com.company;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Scanner;
public class main {

    static Scanner scanner;
    static Scanner fileReader;
    static File f;
    static XSSFWorkbook workbook;
    static XSSFSheet sheet;
    static FileOutputStream output;
    static DecimalFormat df;

    static String file;
    static String output_path;
    static String path;
    static String w_min, w_max;
    static String l_min, l_max;
    static int i_exc, j_exc;

    static String aux;
    static String extension;
    static String[] parts;
    static String fc_low;
    static String fc_high;
    static int quantsPunts;
    static double mitja;
    static int mode;

    public static void main(String[] args) throws IOException {
        try {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Efficiences");
            df = new DecimalFormat("###.###");


            scanner = new Scanner(System.in);
            i_exc = 0;

            Row row;
            Cell cell;

            quantsPunts = 0;

            System.out.println("Welcome to eff calculator. By: Bernat Oller");

            System.out.print("Insert mode (1 = S11 | 2 = Efficiences): ");
            mode = scanner.nextInt();
            scanner.nextLine();
            System.out.println();

            System.out.println("Files must be .txt\n");

            System.out.print("Insert path with all .txt: ");
            path = scanner.nextLine();
            System.out.println();

            System.out.print("Insert path where you want the .xlsx: ");
            output_path = scanner.nextLine();
            System.out.print("Insert name of the .xlsx file: ");
            extension = scanner.nextLine();
            System.out.println();

            output_path = output_path + extension + ".xlsx";

            System.out.println(output_path);

            output = new FileOutputStream(new File(output_path));
            //output = new FileOutputStream("efficiencies_proba.xlsx");

            System.out.print("Insert L min: ");
            l_min = scanner.nextLine();
            System.out.print("Insert L max: ");
            l_max = scanner.nextLine();
            System.out.println();

            System.out.print("Insert W min: ");
            w_min = scanner.nextLine();
            System.out.print("Insert W min: ");
            w_max = scanner.nextLine();
            System.out.println();

            System.out.print("Extension at the end of the file: ");
            extension = scanner.nextLine();
            System.out.println();

            file = l_min + "x" + w_min + extension + ".txt";
            System.out.println("File name: " + file);

            System.out.print("Fc tall baixa: ");
            fc_low = scanner.nextLine();
            System.out.print("Fc tall alta: ");
            fc_high = scanner.nextLine();




            for (int i = Integer.parseInt(l_min); i != Integer.parseInt(l_max) + 10; i = i + 10) {
                row = sheet.createRow(i_exc);
                i_exc++;
                j_exc = 0;
                for (int j = Integer.parseInt(w_min); j != Integer.parseInt(w_max) + 10; j = j + 10) {
                    file = i + "x" + j + extension + ".txt";
                    //System.out.println("File in for: " + file);

                    f = new File(path + file);
                    fileReader = new Scanner(f);

                    quantsPunts = 0;
                    switch (mode){
                        case 1:
                            mitja = -100;
                            fileReader.nextLine();
                            //fileReader.nextLine();
                            break;
                        case 2:
                            mitja = 0;
                            fileReader.nextLine();
                            fileReader.nextLine();
                            break;
                    }
                    while (fileReader.hasNextLine()) {
                        aux = fileReader.nextLine();
                        parts = aux.split("\t");
                        //System.out.println("Part 1: " + Double.parseDouble(parts[0]) + "Part 2: " + Double.parseDouble(parts[1]));

                        if (Double.parseDouble(parts[0]) >= Double.parseDouble(fc_low) && Double.parseDouble(parts[0]) <= Double.parseDouble(fc_high)) {
                            //System.out.println("Dintre rang: " + parts[0]);
                            //System.out.println("Valor eff: " + parts[1]);
                            switch (mode){
                                case 1:
                                    if (Double.parseDouble(parts[1]) > mitja){
                                        mitja = Double.parseDouble(parts[1]);
                                    }
                                    break;
                                case 2:
                                    mitja = mitja + Double.parseDouble(parts[1]);
                                    quantsPunts++;
                                    break;
                            }

                        }
                    }
                    switch (mode){
                        case 1:

                            break;
                        case 2:
                            mitja = mitja / quantsPunts;
                            break;
                    }


                    System.out.print("PCB " + file);
                    System.out.printf(": %f\n", mitja);

                    cell = row.createCell(j_exc);
                    cell.setCellValue(df.format(mitja));
                    j_exc++;


                    fileReader.close();
                    fileReader = null;
                    f = null;
                }

            }

            workbook.write(output);
            output.close();
            System.out.println("Correctly saved");
            System.out.println("Make sure to convert in the excel from string to number.\n");
            System.out.println("Thank you!");

        } catch (IOException e) {
            System.out.println("Could not open file. Close it, check the input values or make sure file " + file + "exists.");
        }
    }
}
