package com.example.homework.hw04;


import com.example.functions.comparingExcelFiles;

public class B11 {
    public static void main(String[] args) {

        String abgabe_directory = "ExcelFiles/Abgabe/Beispiel11_Abgabe.xlsx";
        String loesung_directory = "ExcelFiles/Loesungen/Beispiel11_Loesung.xlsx";

        if (comparingExcelFiles.correctSolution(abgabe_directory,loesung_directory)) {
            System.out.println("Ihre Lösung ist korrekt!");
        }
        else {
            System.out.println("Ihre Lösung ist nicht korrekt!");
            System.out.println(comparingExcelFiles.checkZwischenergebnis(abgabe_directory,loesung_directory));
        }
    }
}