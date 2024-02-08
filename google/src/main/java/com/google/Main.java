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

    private static Sheets sheetsService;
    private static String APPLICATION_NAME = "ChallengeDevTraining";
    private static String spreadSheetID = "1Fqqa-1mtq1OPK6k2MKMjIC70Bqxi49ExEY3Lct4tHo8";

    private static Credential authorize() throws IOException, GeneralSecurityException{
            String credentialPath = "google\\src\\main\\java\\com\\google\\resources\\credentials.json";
            GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JacksonFactory.getDefaultInstance(), new FileReader(credentialPath));

            List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

            GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), 
                JacksonFactory.getDefaultInstance(), 
                clientSecrets, 
                scopes).setDataStoreFactory(new FileDataStoreFactory(new java.io.File("Tokens"))).setAccessType("offline").build();

                Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
                return credential;
    }

    public static Sheets getSheetsService() throws IOException, GeneralSecurityException{
        Credential credential = authorize();
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
         JacksonFactory.getDefaultInstance(),
         credential 
         ).setApplicationName(APPLICATION_NAME).build();
    }


    public static void main(String[] args) throws NumberFormatException, Exception {
        sheetsService = getSheetsService();

        int media;
        String situation;
        int NecessaryToPass;

        int row = 4;

        for(int i=4; i<=27;i++){
            String rangeToAbsence = "C"+row;
            String rangeP1 = "D"+row;
            String rangeP2 = "E"+row;
            String rangeP3 = "F"+row;

            int absence = Integer.parseInt(readCellValue(rangeToAbsence));
            int P1 = Integer.parseInt(readCellValue(rangeP1));
            int P2 = Integer.parseInt(readCellValue(rangeP2));
            int P3 = Integer.parseInt(readCellValue(rangeP3));

            System.out.print(absence + " " + P1 + " " + P2 + " " + P3);

            media = (P1+P2+P3)/3;
            System.out.print(" media: " + media);
            situation = VerificationMedia(media);

            if(absence > 15){
                
                situation = "Reprovado por falta";
                System.out.print(" " + situation);
                NecessaryToPass = 0;
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass) ;

                writeData(row, situation, NecessaryToPass);

            } else if(situation == "Exame Final"){

                System.out.print(" " + situation);
                NecessaryToPass = 100-media;
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass) ;
                
                writeData(row, situation, NecessaryToPass);

            } else if(situation == "Reprovado por nota"){
                System.out.print(" " + situation);
                NecessaryToPass = 0;
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass) ;

                writeData(row, situation, NecessaryToPass);
            } else {
                System.out.print(" " + situation);
                NecessaryToPass = 0;
                System.out.print(" Nota necessaria pra passar é: " + NecessaryToPass) ;

                writeData(row, situation, NecessaryToPass);
            }
            row++;
            System.out.println();
        }

        


        //  PRINT THE SHEET ON TERMINAL
        // if(values == null || values.isEmpty()){
        //     System.out.println("No data");
        // } else {
        //     for(List row : values){
        //         System.out.printf("%s %s %s %s %s\n", row.get(0), row.get(1), row.get(2), row.get(4), row.get(5));
        //     }
        // }


    }

    public static String readCellValue(String range) throws Exception {
        ValueRange response = sheetsService.spreadsheets().values()
                .get(spreadSheetID, range)
                .execute();
        
        List<List<Object>> values = response.getValues();
        
        if (values == null || values.isEmpty()) {
            System.out.println("No data");
            return null;
        } else {
            List<Object> cell = values.get(0);
            if (!cell.isEmpty()) {
                return cell.get(0).toString();
            }
        }
        return null;
    }

    public static String VerificationMedia(int media){
        String disaproved = "Reprovado por nota";
        String avaliation = "Exame Final";
        String aproved = "Aprovado";
        if(media<50){
            return disaproved;
        } else if(media >= 50 && media < 70){
            return avaliation;
        } else if(media >= 70){
            return aproved;
        }
        return null;
    }

    public static void writeData(int row, String situation, int NecessaryToPass) throws IOException{
        
        ValueRange appendBody = new ValueRange()
        .setValues(Arrays.asList(Arrays.asList(situation, NecessaryToPass)));

            sheetsService.spreadsheets().values()
                .append(spreadSheetID, "G"+row+":H"+row, appendBody)
                .setValueInputOption("USER_ENTERED")
                .setInsertDataOption("OVERWRITE")
                .setIncludeValuesInResponse(true)
                .execute();
    }
    
    
}