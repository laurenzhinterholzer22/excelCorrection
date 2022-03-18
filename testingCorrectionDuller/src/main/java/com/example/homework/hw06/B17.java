package com.example.homework.hw06;


import com.example.functions.comparingExcelFiles;

public class B17 {
    public static void main(String[] args) {

        String abgabe_directory = "ExcelFiles/Abgabe/Beispiel17_Abgabe.xlsx";
        String loesung_directory = "ExcelFiles/Loesungen/Beispiel17_Loesung.xlsx";

        if (comparingExcelFiles.correctSolution(abgabe_directory,loesung_directory)) {
            System.out.println("Ihre Lösung ist korrekt!");
        }
        else {
            System.out.println("Ihre Lösung ist nicht korrekt!");
            System.out.println(comparingExcelFiles.checkZwischenergebnis(abgabe_directory,loesung_directory));
        }
    }
}