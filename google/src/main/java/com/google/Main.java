package com.google;

import java.io.FileReader;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;



public class Main {

    private static Sheets sheetsService; //create an object sheetservice to acesss functions to read and write google sheets
    private static String APPLICATION_NAME = "ChallengeDevTraining"; //name of application
    private static String spreadSheetID = "1Fqqa-1mtq1OPK6k2MKMjIC70Bqxi49ExEY3Lct4tHo8"; //id of the sheet



    //Function called authorize to call the credentials and checked if its ok to acess the google sheets
    private static Credential authorize() throws IOException, GeneralSecurityException{
            String credentialPath = "google\\src\\main\\java\\com\\google\\resources\\credentials.json";
            GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JacksonFactory.getDefaultInstance(), new FileReader(credentialPath));

            List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

            GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), 
                JacksonFactory.getDefaultInstance(), 
                clientSecrets, 
                scopes).setDataStoreFactory(new FileDataStoreFactory(new java.io.File("Tokens")))//this file called tokens will store the authorization token
                .setAccessType("offline")
                .build();

                Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user"); 
                return credential;
    }

    //Function called getSheetsService
    public static Sheets getSheetsService() throws IOException, GeneralSecurityException{
        Credential credential = authorize();//calling the function authorize to verify de credential
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), 
         JacksonFactory.getDefaultInstance(),
         credential 
         ).setApplicationName(APPLICATION_NAME).build();
    }

    //Function main, where the alghorithm will run
    public static void main(String[] args) throws NumberFormatException, Exception {
        sheetsService = getSheetsService(); //calling the variable sheetsService and calling the function SheetsService

        int media;  //Variable of the type int to represent the media of an student
        String situation; //Variable of the type String to represent the situation of an student
        int NecessaryToPass; // Variable of the type int to represent the necessary note to pass if the student needs more note to pass

        int row = 4; //Variable to represent the row of the sheets

        for(int i=4; i<=27;i++){ //there will be 24 students, so this process will be done 24 times. start in the row 4 and finish in the row 27
            String rangeToAbsence = "C"+row;
            String rangeP1 = "D"+row;
            String rangeP2 = "E"+row;
            String rangeP3 = "F"+row;
            //Strings to represent the range of each information (Absence, P1, P2, P3) on the sheet.

            int absence = Integer.parseInt(readCellValue(rangeToAbsence));
            int P1 = Integer.parseInt(readCellValue(rangeP1));
            int P2 = Integer.parseInt(readCellValue(rangeP2));
            int P3 = Integer.parseInt(readCellValue(rangeP3));
            //Basically i each variable int, it is calling the readCellValue thats read a value of a cell as a string and convert to int

            System.out.print(absence + " " + P1 + " " + P2 + " " + P3); //print the informations on terminal

            media = (P1+P2+P3)/3; //media of the 3 avaliations
            System.out.print(" media: " + media); //print media on terminal
            situation = VerificationMedia(media); //calling the function VerificationMedia to verify the media and put the situation according to the media.

            if(absence > 15){ //if the absence is above 25% (15 absence)

                situation = "Reprovado por falta";  //situation: repproved by ausence
                System.out.print(" " + situation); //print the situation
                NecessaryToPass = 0; 
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass) ;//print the necessary to pass

                writeData(row, situation, NecessaryToPass); //call the function writeData to write on the sheet

            } else if(situation == "Exame Final"){//if the student is in final exam

                System.out.print(" " + situation);//print the situation
                NecessaryToPass = 100-media;//equation, the result is the necessary to the student pass
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass); // print the necessary to pass
                
                writeData(row, situation, NecessaryToPass); //call the function writeData to write the informations on the sheet

            } else if(situation == "Reprovado por nota"){ //if the student didn't get enough note to be in the exam - repproved by note
                System.out.print(" " + situation);//print the situation
                NecessaryToPass = 0;
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass) ;//print the necessary to pass

                writeData(row, situation, NecessaryToPass); //call the function writeData to write the informations on the sheet
            } else { //if the student got the enough grade to pass - approved
                System.out.print(" " + situation); //print the situation on terminal
                NecessaryToPass = 0;
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass) ; // print the necessary to pass

                writeData(row, situation, NecessaryToPass); //call the function writeData to write the informations on the sheet
            }
            row++; // each time the process is made the row increases one
            System.out.println();//jump one line on terminal
        }
    }

    //Function to read the value of the cell, as parameter the range of the cell
    public static String readCellValue(String range) throws Exception {
        ValueRange response = sheetsService.spreadsheets().values()
                .get(spreadSheetID, range)
                .execute();
        
        List<List<Object>> values = response.getValues();
        
        if (values == null || values.isEmpty()) {
            System.out.println("No data");
            return null;
            //If the cell is empty, print on terminal "no data"
        } else {
            List<Object> cell = values.get(0);
            if (!cell.isEmpty()){
                return cell.get(0).toString();
                //if the cell has information return the information as string
            }
        }
        return null;
    }


    //Function to verify the media and the situation, as parameter the media
    public static String VerificationMedia(int media){
        String disaproved = "Reprovado por nota"; //string to disapproved
        String avaliation = "Exame Final"; //string to final exam
        String aproved = "Aprovado"; // string to approved
        if(media<50){
            return disaproved;
        } else if(media >= 50 && media < 70){
            return avaliation;
        } else if(media >= 70){
            return aproved;
        }
        return null;
    }

    //Function to write on the sheet, the parameters are the row, situation and the necessary to pass
    public static void writeData(int row, String situation, int NecessaryToPass) throws IOException{
        ValueRange appendBody = new ValueRange()
        .setValues(Arrays.asList(Arrays.asList(situation, NecessaryToPass)));

        AppendValuesResponse appendResult = sheetsService.spreadsheets().values()
                .append(spreadSheetID, "G"+row+":H"+row, appendBody) //this fuction has as parameter the id of the sheet, the range of the cells to write data and the values that will write
                .setValueInputOption("USER_ENTERED")
                .setInsertDataOption("OVERWRITE") //this overwrite the cell
                .setIncludeValuesInResponse(true)
                .execute();
    }
}