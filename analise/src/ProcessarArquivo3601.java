import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileOutputStream;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.text.SimpleDateFormat;

public class ProcessarArquivo3601 {

    public Map<String, String[]> processarArquivo(String csvFile) {
        Map<String, String[]> dados3601 = new HashMap<>();        
        String outputExcel = "Login_e_Pausas_das_Operadora.xlsx";

        try {
            // Leitura do CSV
            BufferedReader reader = new BufferedReader(new FileReader(csvFile));
            String line;

            // Mapa para armazenar as informações dos eventos
            Map<String, Map<String, EventDetails>> eventos = new HashMap<>();

            // Ignora o cabeçalho
            reader.readLine();

            while ((line = reader.readLine()) != null) {
                String[] columns = line.split(";");
                String agente = columns[0].replace("\"", "");
                String tipoEvento = columns[4].replace("\"", "");
                String loginInicial = columns[1].replace("\"", "");
                String loginFinal = columns[2].replace("\"", "");

                // Ignora eventos sem tipo, login inicial ou login final
                if (tipoEvento.isEmpty() || loginInicial.isEmpty() || loginFinal.isEmpty()) continue;

                // Adiciona o evento no mapa
                if (!eventos.containsKey(agente)) {
                    eventos.put(agente, new HashMap<>());
                }

                if (!eventos.get(agente).containsKey(tipoEvento)) {
                    eventos.get(agente).put(tipoEvento, new EventDetails());
                }

                EventDetails eventDetails = eventos.get(agente).get(tipoEvento);
                eventDetails.incrementarQuantidade();

                // Calcula a duração
                long duracaoSegundos = timeToSeconds(loginFinal) - timeToSeconds(loginInicial);
                eventDetails.incrementarDuracao(duracaoSegundos);

                // Calcula o menor login inicial e o maior login final
                if (eventDetails.getLoginInicial() == null || timeToSeconds(loginInicial) < timeToSeconds(eventDetails.getLoginInicial())) {
                    eventDetails.setLoginInicial(loginInicial);
                }

                if (eventDetails.getLoginFinal() == null || timeToSeconds(loginFinal) > timeToSeconds(eventDetails.getLoginFinal())) {
                    eventDetails.setLoginFinal(loginFinal);
                }               

            }
            reader.close();

            // Criação do Excel
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Analise_Operação");

            // Cabeçalho da planilha
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Agente");
            headerRow.createCell(1).setCellValue("LI");
            headerRow.createCell(2).setCellValue("LF");
            headerRow.createCell(3).setCellValue("TL");
            headerRow.createCell(4).setCellValue("Break");
            headerRow.createCell(5).setCellValue("Toilet");
            headerRow.createCell(6).setCellValue("Lanche");
            headerRow.createCell(7).setCellValue("Ginástica");
            headerRow.createCell(8).setCellValue("Assuntos Internos");
            headerRow.createCell(9).setCellValue("Outros");

            // Criar um TreeMap com comparador personalizado para ordenar numericamente
            Map<String, Map<String, EventDetails>> eventosOrdenados = new TreeMap<>(Comparator.comparingInt(s -> Integer.parseInt(s.replace("ope", ""))));

            // Adicionar os dados ao novo TreeMap ordenado
            eventosOrdenados.putAll(eventos);

            // Preenchendo os dados
            int rowNum = 1;
            for (Map.Entry<String, Map<String, EventDetails>> entryAgente : eventosOrdenados.entrySet()) {
                String agente = entryAgente.getKey();
                Map<String, EventDetails> detalhesEventos = entryAgente.getValue();

                // Encontrar as pausas
                Map<String, EventDetails> pausas = new HashMap<>();

                for (Map.Entry<String, EventDetails> entryEvento : detalhesEventos.entrySet()) {
                    String tipoEvento = entryEvento.getKey();
                    EventDetails eventDetails = entryEvento.getValue();
                    if (isPausa(tipoEvento)) {
                        pausas.put(tipoEvento, eventDetails);
                    }
                }
                String tlFormatado = secondsToTime(timeToSeconds(getLoginFinal(detalhesEventos)) - timeToSeconds(getLoginInicial(detalhesEventos)));

                // Armazena LI, LF e TL no mapa
                dados3601.put(agente, new String[]{getLoginInicial(detalhesEventos), getLoginFinal(detalhesEventos),
                            tlFormatado,formatPausa(pausas.get("Break")),formatPausa(pausas.get("Toilet")),
                            formatPausa(pausas.get("Lanche")),formatPausa(pausas.get("Ginástica")),
                            formatPausa(pausas.get("Assuntos Internos")),formatPausa(pausas.get("Outros"))});

                //System.out.println("Dados agente \n"+ dados3601);
                // Criando a linha na planilha
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(agente);  // Agente
                row.createCell(1).setCellValue(getLoginInicial(detalhesEventos));  // LI
                row.createCell(2).setCellValue(getLoginFinal(detalhesEventos));  // LF
                row.createCell(3).setCellValue(secondsToTime(timeToSeconds(getLoginFinal(detalhesEventos)) - timeToSeconds(getLoginInicial(detalhesEventos))));  // TL
                row.createCell(4).setCellValue(formatPausa(pausas.get("Break")));  // Break
                row.createCell(5).setCellValue(formatPausa(pausas.get("Toilet")));  // Toilet
                row.createCell(6).setCellValue(formatPausa(pausas.get("Lanche")));  // Lanche
                row.createCell(7).setCellValue(formatPausa(pausas.get("Ginástica")));  // Ginástica
                row.createCell(8).setCellValue(formatPausa(pausas.get("Assuntos Internos")));  // Assuntos Internos
                row.createCell(9).setCellValue(formatPausa(pausas.get("Outros")));  // Outros
            }   

            // Escrevendo o arquivo Excel
            FileOutputStream fileOut = new FileOutputStream(outputExcel);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
             
            System.out.println("Arquivo Excel gerado com sucesso!");
            //JOptionPane.showMessageDialog(null, "Arquivo gerado com sucesso! \n Consulte sua pasta");

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Erro ao processar o arquivo: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
        }        

        return dados3601;
        
    }

    // Método para abrir o seletor de arquivos
    private static String selectFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Selecione um arquivo CSV");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Arquivos CSV", "csv"));

        int result = fileChooser.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            return fileChooser.getSelectedFile().getAbsolutePath();
        } else {
            return null;
        }
    }

    // Converte a string de tempo (hh:mm:ss ou yyyy-MM-dd HH:mm) para segundos
    private static long timeToSeconds(String time) {
        try {
            // Verifica se o valor é uma data (yyyy-MM-dd HH:mm)
            if (time.contains("-") && time.contains(":")) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                return (long) (sdf.parse(time).getTime() / 1000); // Converte para segundos
            } else {
                // Formato padrão de tempo (hh:mm:ss)
                String[] parts = time.split(":");
                long horas = Integer.parseInt(parts[0]);
                long minutos = Integer.parseInt(parts[1]);
                long segundos = Integer.parseInt(parts[2]);
                return horas * 3600 + minutos * 60 + segundos;
            }
        } catch (Exception e) {
            System.err.println("Erro ao converter tempo: " + time);
            return 0;
        }
    }

    // Converte o tempo em segundos para o formato hh:mm:ss
    private static String secondsToTime(long seconds) {
        long horas = seconds / 3600;
        long minutos = (seconds % 3600) / 60;
        long segundos = seconds % 60;
        return String.format("%02d:%02d:%02d", horas, minutos, segundos);
    }

    // Verifica se o evento é uma pausa
    private static boolean isPausa(String tipoEvento) {
        return 
            tipoEvento.equalsIgnoreCase("Break") ||
            tipoEvento.equalsIgnoreCase("Lanche") ||
            tipoEvento.equalsIgnoreCase("Ginastica")||
            tipoEvento.equalsIgnoreCase("Toilet") ||
            tipoEvento.equalsIgnoreCase("Assuntos Internos");
    }

    // Obtém o Login Inicial (menor hora)
    private static String getLoginInicial(Map<String, EventDetails> eventos) {
        String menorHora = null;
        for (Map.Entry<String, EventDetails> entry : eventos.entrySet()) {
            String menorHoraInicial = entry.getValue().getLoginInicial();
            String loginInicial = menorHoraInicial.split(" ")[1];

            if (menorHora == null || timeToSeconds(loginInicial) < timeToSeconds(menorHora)) {
                menorHora = loginInicial;
            }
        }
        return menorHora;
    }

    // Obtém o Login Final (maior hora)
    private static String getLoginFinal(Map<String, EventDetails> eventos) {
        String maiorHora = null;
        for (Map.Entry<String, EventDetails> entry : eventos.entrySet()) {
            String maiorHoraFinal = entry.getValue().getLoginFinal();
            String loginFinal = maiorHoraFinal.split(" ")[1];
            if (maiorHora == null || timeToSeconds(loginFinal) > timeToSeconds(maiorHora)) {
                maiorHora = loginFinal;
            }
        }
        return maiorHora;
    }  

    private static String formatPausa(EventDetails eventDetails) {
        if (eventDetails == null) {
            return "0 - 00:00:00";
        }
        long duracaoSegundos = eventDetails.getDuracao();
        String duracaoFormatada = secondsToTime(duracaoSegundos);
        return eventDetails.getQuantidade() + " - " + duracaoFormatada;
    }
}