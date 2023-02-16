package com.mendes.controleponto;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import java.io.IOException;
import java.security.GeneralSecurityException;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Optional;

public class ControlePonto {

    // Configurações de autenticação
    private static final String APPLICATION_NAME = "Controle de Ponto Mendes";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    // Colunas de horas
    private static final String[] COLUNAS_HORAS = {"D", "E", "F", "G", "H", "I"};
    // Linha inicial para registro de ponto na planilha
    private static final int LINHA_INICIAL = 7;
    // Mensagens de erro
    private static final String DATA_NAO_ENCONTRADA = "Não foi encontrada a data atual na planilha";
    private static final String SEM_HORARIO_LIVRE = "Não foi encontrado horário livre para registrar!";
    private static final String SEM_PLANILHA = "Não foi encontrada planilha no formato esperado!";
    // Mensagens informativas
    private static final String CELULAS_ATUALIZADAS = "%d células atualizadas na planilha\n";
    // Padrões de data/hora
    private static final String PADRAO_DATA = "dd/MM/yyyy";
    private static final String PADRAO_HORA = "HH:mm";
    // Nome da aba da planilha com as horas trabalhadas
    private static final String NOME_ABA = "Horas!";
    // Range com a coluna de data
    private static final String RANGE_DATA = "Horas!A:A";
    // Determina que os valores vão ser interpretados como se estivessem sido digitados diretamente na planilha.
    private static final String INPUT_OPTION = "USER_ENTERED";
    // Constantes relacionadas a autenticação do usuário
    private static final String USER_ID = "user";
    private static final String ACESS_TYPE = "offline";
    private static final String AUTH_TOKEN = "/client_secret.json";
    private static final String TOKENS_DIR = String.format("%s/.cpmTokens", System.getProperty("user.home"));
    // Query para busca no Drive
    private static final String QUERY_BUSCA_DRIVE = "name contains 'DEMONSTRATIVO DAS HORAS CONTRATADAS' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";

    public static void main(String... args) throws IOException, GeneralSecurityException, Exception {
        final NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();
        Credential credential = authorize(httpTransport);
        Sheets sheetsService = buildSheetsService(httpTransport, credential);
        Drive driveService = buildDriveService(httpTransport, credential);
        String idPlanilha = retornaPlanilhaComNomePadrao(driveService);
        if (idPlanilha.isEmpty()) {
            System.out.println(SEM_PLANILHA);
            return;
        }
        int linha = buscaLinha(sheetsService, idPlanilha);
        if (linha < LINHA_INICIAL) {
            System.out.println(DATA_NAO_ENCONTRADA);
            return;
        }
        String coluna = buscaColuna(linha, sheetsService, idPlanilha);
        if (coluna.isEmpty()) {
            System.out.println(SEM_HORARIO_LIVRE);
            return;
        }
        inserePonto(coluna, linha, sheetsService, idPlanilha);
    }

    /**
     * Retorna a planilha com o nome padrão definido por QUERY_BUSCA_DRIVE
     * 
     * @param driveService
     * @return ID da planilha retornada
     * @throws IOException 
     */
    public static String retornaPlanilhaComNomePadrao(Drive driveService) throws IOException {
        Optional<File> file = driveService.files().list().setQ(QUERY_BUSCA_DRIVE).execute().getFiles().stream().findFirst();
        if (file.isEmpty()) {
            return new String();
        }
        return file.get().getId();
    }

    /**
     * Insere registro de ponto na planilha do Sheets. Será registrado na coluna
     * e linha passada por parâmetro.
     *
     * @param coluna
     * @param linha
     * @param sheetsService
     * @param idPlanilha
     * @throws IOException
     */
    private static void inserePonto(String coluna, int linha, Sheets sheetsService, String idPlanilha) throws IOException {
        SimpleDateFormat horaFormat = new SimpleDateFormat(PADRAO_HORA);
        String horaAtual = horaFormat.format(new Date());
        String rangeHoraLivre = NOME_ABA + coluna + linha;
        List<List<Object>> valuesHoraLivre = Arrays.asList(Arrays.asList(horaAtual));
        ValueRange requestBodyHoraLivre = new ValueRange().setValues(valuesHoraLivre);
        UpdateValuesResponse response = sheetsService.spreadsheets().values()
                .update(idPlanilha, rangeHoraLivre, requestBodyHoraLivre)
                .setValueInputOption(INPUT_OPTION)
                .execute();
        System.out.printf(CELULAS_ATUALIZADAS, response.getUpdatedCells());
    }

    /**
     * Busca primeira coluna de hora livre na linha especificada
     *
     * @param linha
     * @param sheetsService
     * @param idPlanilha 
     * @return Coluna de hora livre
     * @throws IOException
     */
    private static String buscaColuna(int linha, Sheets sheetsService, String idPlanilha) throws IOException {
        for (String colunaHora : COLUNAS_HORAS) {
            String rangeHora = NOME_ABA + colunaHora + linha;
            ValueRange responseHora = sheetsService.spreadsheets().values().get(idPlanilha, rangeHora)
                    .execute();
            List<List<Object>> valuesHora = responseHora.getValues();
            if (valuesHora == null || valuesHora.isEmpty() || valuesHora.get(0).get(0).toString().isEmpty()) {
                return colunaHora;
            }
        }
        return new String();
    }

    /**
     * Retorna o número da linha correspondente a data atual.
     *
     * @param sheetsService
     * @param idPlanilha 
     * @return Número da linha correspondente a data atual
     * @throws IOException
     */
    private static int buscaLinha(Sheets sheetsService, String idPlanilha) throws IOException {
        SimpleDateFormat dateFormat = new SimpleDateFormat(PADRAO_DATA);
        String dataAtual = dateFormat.format(new Date());
        String range = RANGE_DATA;
        ValueRange response = sheetsService.spreadsheets().values().get(idPlanilha, range)
                .execute();
        List<List<Object>> values = response.getValues();
        int rowIndex = LINHA_INICIAL - 1;
        for (int i = 0; i < values.size(); i++) {
            List<Object> row = values.get(i);
            if (!row.isEmpty() && row.get(0).toString().equals(dataAtual)) {
                rowIndex = i + 1;
                break;
            }
        }

        return rowIndex;
    }

    /**
     * Configura o serviço do Sheets responsável por operações em planilhas.
     *
     * @param httpTransport
     * @param credential
     * @return Serviço do Sheets configurado
     * @throws Exception
     */
    public static Sheets buildSheetsService(NetHttpTransport httpTransport, Credential credential) throws Exception {
        return new Sheets.Builder(httpTransport, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    /**
     * Configura o serviço do Drive responsável por operações em arquivos.
     *
     * @param httpTransport
     * @param credential
     * @return Serviço do Drive configurado
     * @throws Exception
     */
    public static Drive buildDriveService(NetHttpTransport httpTransport, Credential credential) throws Exception {
        return new Drive.Builder(httpTransport, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    /**
     * Retorna autorização do usuário para acesso as planilhas.
     *
     * @param httpTransport
     * @return Autorização do usuário para acesso as planilhas
     * @throws Exception
     */
    private static Credential authorize(NetHttpTransport httpTransport) throws Exception {
        InputStream in = ControlePonto.class.getResourceAsStream(AUTH_TOKEN);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                httpTransport, JSON_FACTORY, clientSecrets, List.of(SheetsScopes.SPREADSHEETS, DriveScopes.DRIVE_READONLY))
                .setAccessType(ACESS_TYPE)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIR)))
                .build();
        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize(USER_ID);
        return credential;
    }

}
