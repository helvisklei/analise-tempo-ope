import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.io.FileOutputStream;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.Normalizer;
import java.text.SimpleDateFormat;

public class ProcessarArquivo3601 {

    public Map<String, String[]> processarArquivo(String csvFile) {
        Map<String, String[]> dados3601 = new HashMap<>();        
        String outputExcel = "Login_e_Pausas_das_Operadora.xlsx"; 

        try {
            // Leitura do CSV
            BufferedReader reader = new BufferedReader(new InputStreamReader( new FileInputStream(csvFile),StandardCharsets.UTF_8)
                );
            String line;

            // Mapa para armazenar as informações dos eventos
            Map<String, Map<String, EventDetails>> eventos = new HashMap<>();

            // Ignora o cabeçalho
            reader.readLine();

            while ((line = reader.readLine()) != null) {
                String[] columns = line.split(";");
                String agente = columns[0].replace("\"", "").trim();
                String tipoEvento = columns[4].replace("\"", "").trim();
                String loginInicial = columns[1].replace("\"", "").trim();
                String loginFinal = columns[2].replace("\"", "").trim();

                // Ignora eventos sem tipo, login inicial ou login final
                if (tipoEvento.isEmpty() || loginInicial.isEmpty() || loginFinal.isEmpty()) continue;

                // Adiciona o evento no mapa
                if (!eventos.containsKey(agente)) {
                    eventos.put(agente, new HashMap<>());
                    //System.out.println(agente);
                }

                if (!eventos.get(agente).containsKey(tipoEvento)) {
                    eventos.get(agente).put(tipoEvento.trim(), new EventDetails());
                   // System.out.println("Tipo evento no if "+ tipoEvento.trim());// localiza a ginastica dento tipo evento
                    //System.out.println(agente +" Operadora cm evernto " + eventos.get(agente));
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
            Map<String, Map<String, EventDetails>> eventosOrdenados = new TreeMap<>(Comparator.comparingInt(s -> Integer.parseInt(s.replace("ope", "").trim())));// coloquei trim

            // Adicionar os dados ao novo TreeMap ordenado
            eventosOrdenados.putAll(eventos);

            // Preenchendo os dados
            int rowNum = 1;
            for (Map.Entry<String, Map<String, EventDetails>> entryAgente : eventosOrdenados.entrySet()) {
                String agente = entryAgente.getKey();
                Map<String, EventDetails> detalhesEventos = entryAgente.getValue();
                //System.out.println(" detalhe evento em map " + detalhesEventos);// verificar se a pausa ginastica é armazenada

                // Encontrar as pausas
                Map<String, EventDetails> pausas = new HashMap<>(); // vazio
                EventDetails outrosEventos = new EventDetails(); // Cria um objeto para "Outros"

                //System.out.println("pausas " + pausas);

                for (Map.Entry<String, EventDetails> entryEvento : detalhesEventos.entrySet()) {
                    String tipoEvento = entryEvento.getKey();
                    EventDetails eventDetails = entryEvento.getValue();
                    if (isPausa(tipoEvento)) {                       
                        pausas.put(tipoEvento, eventDetails);
                        //System.out.println("pausas " + pausas);// analisando as inclusões das pausas no mapa
                    }else if (!isCategoriaEspecifica(tipoEvento)){
                        // Se não for pausa predefinidas diciona em outros
                        outrosEventos.incrementarQuantidade();
                        outrosEventos.incrementarDuracao(eventDetails.getDuracao());
                    }                    
                }

                            // Adiciona "Outros" ao mapa de pausas (se houver eventos)
                if (outrosEventos.getQuantidade() > 0) {
                    pausas.put("Outros", outrosEventos);
                }
                String tlFormatado = secondsToTime(timeToSeconds(getLoginFinal(detalhesEventos)) - timeToSeconds(getLoginInicial(detalhesEventos)));
                // Calcula os totais em segundos
                Long breakSegundos = timeToSeconds(extrairDuracao(formatPausa(pausas.get("Break"))));
                Long toiletSegundos = timeToSeconds(extrairDuracao(formatPausa(pausas.get("Toilet")))); 
                Long lancheSegundos = timeToSeconds(extrairDuracao(formatPausa(pausas.get("Lanche"))));
                Long ginasticaSegundos = timeToSeconds(extrairDuracao(formatPausa(pausas.get("Ginástica"))));
                Long assuntosInternosSegundos =timeToSeconds(extrairDuracao(formatPausa(pausas.get("Assuntos Internos"))));
                Long outrosSegundos =timeToSeconds(extrairDuracao(formatPausa(pausas.get("Outros"))));
                // Soma os totais em segundos
                long totalDePausas = breakSegundos + toiletSegundos+lancheSegundos+ginasticaSegundos+assuntosInternosSegundos+outrosSegundos; 

                String  totalDePausasFormatado  = secondsToTime(totalDePausas);
                // Armazena LI, LF e TL no mapa
                dados3601.put(agente, new String[]{getLoginInicial(detalhesEventos), getLoginFinal(detalhesEventos),
                    tlFormatado,formatPausa(pausas.get("Break")),formatPausa(pausas.get("Toilet")),
                    formatPausa(pausas.get("Lanche")),formatPausa(pausas.get("Ginástica")),
                    formatPausa(pausas.get("Assuntos Internos")),formatPausa(pausas.get("Outros")),
                    totalDePausasFormatado}); 
               
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
            tipoEvento.trim().equalsIgnoreCase("Lanche") ||
            tipoEvento.trim().equalsIgnoreCase("Ginástica")||
            tipoEvento.trim().equalsIgnoreCase("Toilet") ||
            tipoEvento.trim().equalsIgnoreCase("Assuntos Internos");
    }

    // Verifica se o evento é uma categoria específica
    private static boolean isCategoriaEspecifica(String tipoEvento) {
        return CATEGORIAS_ESPECIFICAS.contains(tipoEvento.trim());
    }

    // cria uma categoria que não deve ser lida e contabilizada
    private static final Set<String> CATEGORIAS_ESPECIFICAS = Set.of(
        "Pausa Sistema", "Login", "Pré-pausa", "Pré-pausa: Break", "Pré-pausa: Lanche", "Pré-pausa: Toilet",
         "Pré-pausa: Login", "Pré-pausa: Assuntos Internos"
    );

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
    /**
     * Extrai a duração formatada de uma string no formato "quantidade - duração".
     *
     * @param pausaFormatada String no formato "quantidade - duração".
     * @return String contendo apenas a duração no formato "HH:mm:ss".
    */
    private static String extrairDuracao(String pausaFormatada) {
        if (pausaFormatada == null || !pausaFormatada.contains(" - ")) {
            return "00:00:00"; // Retorna um valor padrão caso o formato seja inválido
        }
        // Encontra o índice do hífen e adiciona 3 para pular o espaço, hífen e espaço
        int indiceHifen = pausaFormatada.indexOf(" - ") + 3;
        // Retorna a substring a partir do índice do hífen até o final da string
        return pausaFormatada.substring(indiceHifen);
    }

    public class StringUtils {
        public static String removerAcentos(String str) {
            String normalized = Normalizer.normalize(str, Normalizer.Form.NFD);
            Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
            return pattern.matcher(normalized).replaceAll("");
        }
    }

}