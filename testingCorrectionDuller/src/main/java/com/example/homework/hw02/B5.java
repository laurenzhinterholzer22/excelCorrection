package com.example.homework.hw02;


import com.example.functions.comparingExcelFiles;

public class B5 {
    public static void main(String[] args) {

        String abgabe_directory = "ExcelFiles/Abgabe/Beispiel05_Abgabe.xlsx";
        String loesung_directory = "ExcelFiles/Loesungen/Beispiel05_Loesung.xlsx";

        if (comparingExcelFiles.correctSolution(abgabe_directory,loesung_directory)) {
            System.out.println("Ihre Lösung ist korrekt!");
        }
        else {
            System.out.println("Ihre Lösung ist nicht korrekt!");
            System.out.println(comparingExcelFiles.checkZwischenergebnis(abgabe_directory,loesung_directory));
            //System.out.println(comparingExcelFiles.wrongCell(abgabe_directory,loesung_directory));
        }
    }
}