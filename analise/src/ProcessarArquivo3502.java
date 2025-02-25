import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ProcessarArquivo3502 {
    public static final String THPA = "05:20:20"; // Total de horas trabalhadas na PA
    public static final String TTL = "03:10:00"; // Tempo total em ligação
    public static void processarDados(String csvFile,  Map<String, String[]> dados3601) {
        //Map<String, List<String[]>> dadosPorAgente = new HashMap<>();
        // Usando TreeMap com um Comparator personalizado
        Map<String, List<String[]>> dadosPorAgente = new TreeMap<>(Comparator.comparingInt(ProcessarArquivo3502::extrairNumeroAgente));
        long tempoTrabalhadoTotal = 0;
        long tempoTotalTma = 0;
        int tempoTotalRepeticoes = 0;
        long porcentagemTotalHorasTrabalhadasNPa = 0;
        long porcentagemTempoTotalEmLigacao = 0;
        long totalPorcentoThpa = 0;
        long totalPorcentoTtl = 0;
        long totalMtsl = 0;

        try (BufferedReader br = Files.newBufferedReader(Paths.get(csvFile));
                XSSFWorkbook workbook = new XSSFWorkbook()) { // Usando XSSFWorkbook para criar um arquivo Excel
            
                    
            //Map<String, CellStyle> estilos = EstilosExcel.criarEstilos(workbook);         
            // Obter a data atual
            Date dataAtual = new Date(); 
            // Estilo para a linha de totalização
            CellStyle totalStyle = EstilosExcel.getTotalCellStyle(workbook);

            // Formatar a data para um formato de string (exemplo: "yyyy-MM-dd")
            SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
            String nomePlanilha = "Análise das OPE - " + sdf.format(dataAtual);       

            Sheet sheet = workbook.createSheet(nomePlanilha);

            // borda
            Row headerRow = sheet.createRow(0);

            // Cabeçalhos
            String[] headers = { "DATA", "OPERADORA","LI","LF", "1ºL", "TULD","TL","JTD", "QTD", "%THPA", "TTL", "%TTL", "TMA", "MTSL"
                                ,"","BREAK","TOILET","LANCHE","GINÁSTICA","ASS. INTERNO", "OUTROS"}; 
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(EstilosExcel.getHeaderCellStyle(workbook)); // Estilo para o cabeçalho
                //sheet.autoSizeColumn(i);
            }

            String linha;
            br.readLine(); // Pular cabeçalho
            int rowIndex = 1;
            int maxColumns = 0;

            while ((linha = br.readLine()) != null) {
                String[] colunas = linha.split(";");
                for (int i = 1; i <= 2; i++) { // Remover aspas das colunas 2 e 9
                    colunas[i] = colunas[i].replace("\"", "");
                }
                colunas[9] = colunas[9].replace("\"", "");
                String agente = colunas[2];

                dadosPorAgente.computeIfAbsent(agente, k -> new ArrayList<>()).add(colunas);
            }

            for (Map.Entry<String, List<String[]>> entry : dadosPorAgente.entrySet()) {
                String agente = entry.getKey();
                List<String[]> registros = entry.getValue();

                String data = registros.get(0)[1].split(" ")[0];
                String menorHoraInicial = registros.stream()
                        .map(r -> r[1].split(" ")[1])
                        .min(String::compareTo)
                        .orElse("");

                String maiorHoraFinal = registros.stream()
                        .map(r -> r[9].split(" ")[1])
                        .max(String::compareTo)
                        .orElse("");

                int repeticoes = registros.size();
                tempoTotalRepeticoes += repeticoes; // QTD - quantidade total diaria
                //System.out.println(" linha 93 - Total " + tempoTotalRepeticoes + " = " + registros.size()); // analisando codigo

                long tempoTrabalhado = registros.stream()
                        .mapToLong(r -> calcularDiferencaSegundos(r[1], r[9]))
                        .sum(); // retorna a soma das diferença TTL

                long mtsl = registros.stream()
                        .mapToLong(r -> calcularDiferencaSegundos(r[1], r[9]))
                        .max()
                        .orElse(0L); // Retorna maior tempo MTSL

                tempoTrabalhadoTotal += tempoTrabalhado;
                String tma = formatarSegundos(tempoTrabalhado / repeticoes);
                tempoTotalTma += (converterParaSegundos(tma)); // tempo total TMA

                porcentagemTotalHorasTrabalhadasNPa = calcularPorcentagem(tempoTrabalhado, converterParaSegundos(THPA)); // %THPA
                porcentagemTempoTotalEmLigacao = calcularPorcentagem(tempoTrabalhado, converterParaSegundos(TTL)); // %TTL
                totalMtsl += mtsl;
                totalPorcentoThpa += porcentagemTotalHorasTrabalhadasNPa;
                totalPorcentoTtl += porcentagemTempoTotalEmLigacao;

                // Criando uma nova linha para o agente no arquivo Excel
                Row row = sheet.createRow(rowIndex++);
                CellStyle estiloBorda = EstilosExcel.getCellStyleWithBorders(workbook);

                for (int i = 0; i <= 20; i++) { // demarca as linhas com as bordas
                    Cell cell = row.createCell(i);                    

                    // Definir valores conforme a coluna
                    switch (i) {
                        case 0 -> 
                        {   
                            // Crie um SimpleDateFormat para analisar a string
                            SimpleDateFormat parser = new SimpleDateFormat("yyyy-MM-dd");
                            
                            try {
                                // Converta a string em um objeto Date
                                Date dataConvertida = parser.parse(data);
                                
                                // Crie um SimpleDateFormat para formatar a data no formato desejado
                                SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
                                
                                // Formate a data e defina o valor na célula
                                String dataFormatada = formatter.format(dataConvertida);
                                cell.setCellValue(dataFormatada);
                            } catch (ParseException e) {
                                e.printStackTrace();
                                // Trate a exceção de parsing aqui
                            }
                           // cell.setCellValue(sdf.parse(data.toString()));// DATA sdf.format(dataAtual)
                        }
                        case 1 -> cell.setCellValue(agente);// OPE
                        //case 2 -> cell.setCellValue("");// LI
                        //case 3 -> cell.setCellValue("");// LF
                        case 4 -> cell.setCellValue(menorHoraInicial);// era posição 2 - 1 L
                        case 5 -> cell.setCellValue(maiorHoraFinal); // era posição 3 - TULD
                        //case 6 -> cell.setCellValue("");// TL  segunda planilha dados vindo abaixo
                        //case 7 -> cell.setCellValue("");// era posição 4 QTD
                        case 8 -> cell.setCellValue(repeticoes);// era posição 4 QTD
                        case 9 -> cell.setCellValue(porcentagemTotalHorasTrabalhadasNPa + "%");// ra posição 5 QTD
                        case 10 -> cell.setCellValue(formatarSegundos((tempoTrabalhado)));// ra posição ttl
                        case 11 -> cell.setCellValue(porcentagemTempoTotalEmLigacao + "%");
                        case 12 -> cell.setCellValue(tma);
                        case 13 -> cell.setCellValue(formatarSegundos(mtsl));                      
                      
                    }
                  
                    if (i == 4) {
                        CellStyle estiloCondicional = EstilosExcel.getConditionalCellStyle(workbook, menorHoraInicial);
                        CellStyle bordaColunaMenorHora = EstilosExcel.aplicarBordas(estiloCondicional); 
                                               
                        cell.setCellStyle(bordaColunaMenorHora); // Aplicar estilo com borda e cor

                    } else {
                        cell.setCellStyle(estiloBorda); // Aplicar apenas a borda normal
                    }
                } 
                int lastCellNum = row.getLastCellNum();
                
                if (lastCellNum > maxColumns) {
                    maxColumns = lastCellNum;
                }

                 //  Inserir dados do 3601 entre as colunas
                 if (dados3601.containsKey(agente)) {
                    String[] dadosAgente3601 = dados3601.get(agente);
                    row.createCell(2).setCellValue(dadosAgente3601[0]); // LI
                    row.createCell(3).setCellValue(dadosAgente3601[1]); // LF
                    row.createCell(6).setCellValue(dadosAgente3601[2]); // TL    
                    row.createCell(7).setCellValue(formatarSegundos(converterParaSegundos(dadosAgente3601[2])- converterParaSegundos(dadosAgente3601[9]))); // JTD  
                   //row.createCell(14).setCellValue(""); // Break                                  
                    row.createCell(15).setCellValue(dadosAgente3601[3]); // Break
                    row.createCell(16).setCellValue(dadosAgente3601[4]); // Toilet
                    row.createCell(17).setCellValue(dadosAgente3601[5]); // Lanche
                    row.createCell(18).setCellValue(dadosAgente3601[6]); // Ginastica
                    row.createCell(19).setCellValue(dadosAgente3601[7]); // Assustos internos
                    row.createCell(20).setCellValue(dadosAgente3601[8]); // Outros

                    

                    row.getCell(0).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(0);

                    row.getCell(1).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(1);

                    row.getCell(2).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(2);

                    row.getCell(3).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(3);

                    row.getCell(5).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(5);

                    row.getCell(6).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(6);

                    row.getCell(7).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(7);

                    row.getCell(8).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(8);

                    row.getCell(9).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(9);  
                    
                    //row.getCell(10).setCellStyle(estiloBorda);
                    //sheet.autoSizeColumn(10);

                    row.getCell(11).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(11);

                    row.getCell(12).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(12);

                    row.getCell(13).setCellStyle(estiloBorda); 
                    sheet.autoSizeColumn(13);

                    row.getCell(15).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(15);

                    row.getCell(16).setCellStyle(estiloBorda); 
                    sheet.autoSizeColumn(16);

                    row.getCell(17).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(17);

                    row.getCell(18).setCellStyle(estiloBorda); 
                    sheet.autoSizeColumn(18);

                    row.getCell(19).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(19);

                    row.getCell(20).setCellStyle(estiloBorda);
                    sheet.autoSizeColumn(20);
                    //row.getCell(20).setCellStyle(estiloBorda);   
                    //sheet.autoSizeColumn(20);                                     
                    
                }   
                //System.out.println("Total de coolnas " + maxColumns);             
                
            }

            // Adicionando a última linha com os totais
            Row totalRow = sheet.createRow(rowIndex + 1);    
            
             // 🔹 Mesclar a célula inicial nas linhas 1 e 2
             int ultimaLinha = sheet.getLastRowNum();      
            
            totalRow.createCell(0).setCellValue("Total OPE");         
            totalRow.createCell(1).setCellValue(dadosPorAgente.size());
            // Mescla as células da coluna A nas linhas 36 e 37
            CellRangeAddress cellRangeAddress = new CellRangeAddress(ultimaLinha, sheet.getLastRowNum()+1, 0, 0);
            sheet.addMergedRegion(cellRangeAddress);
            // Mescla as células da coluna B nas linhas 36 e 37
            CellRangeAddress cellRangeAddress2 = new CellRangeAddress(ultimaLinha, sheet.getLastRowNum()+1, 1, 1);
            sheet.addMergedRegion(cellRangeAddress2);
            
            totalRow.createCell(2).setCellValue("Total Geral"); // Total de agentes
            totalRow.createCell(3).setCellValue("");
            totalRow.createCell(4).setCellValue("");
            totalRow.createCell(5).setCellValue("");
            totalRow.createCell(6).setCellValue("");
            totalRow.createCell(7).setCellValue("");// 
            totalRow.createCell(8).setCellValue(tempoTotalRepeticoes);// QTD - tempo total repetições QTD
            totalRow.createCell(9).setCellValue(totalPorcentoThpa); // %THPA
            totalRow.createCell(10).setCellValue(formatarSegundos(tempoTrabalhadoTotal).trim()); // TTL Total de tempo trabalhado formatado
            totalRow.createCell(11).setCellValue(totalPorcentoTtl); // %TTL
            totalRow.createCell(12).setCellValue(formatarSegundos(tempoTotalTma).trim());// total de tempo TMA
            totalRow.createCell(13).setCellValue(formatarSegundos(totalMtsl).trim());// MTSL

            // MEscla a celula total
            sheet.addMergedRegion(new CellRangeAddress(ultimaLinha, sheet.getLastRowNum(), 2, 7)); // mescla celula em total
                       
            
            // 🔹 Mesclar a célula media
            sheet.addMergedRegion(new CellRangeAddress(ultimaLinha + 1, ultimaLinha + 1, 2, 7)); // mescla celula em Média
            for (int i = 0; i <= 15; i++) {
                if (totalRow.getCell(i) != null) {
                    totalRow.getCell(i).setCellStyle(totalStyle);// coloca estilo azul em total  até 14                    
                }
            }

            Row totalRow2 = sheet.createRow(sheet.getLastRowNum() + 1);
            totalRow2.createCell(0).setCellValue("");
            totalRow2.createCell(1).setCellValue(" ");
            totalRow2.createCell(2).setCellValue("Média Geral");
            totalRow2.createCell(3).setCellValue("");
            totalRow2.createCell(4).setCellValue(" ");
            totalRow2.createCell(5).setCellValue(" ");
            //totalRow2.createCell(6).setCellValue(" ");
            totalRow2.createCell(7).setCellValue(Math.round(tempoTotalRepeticoes / (double) dadosPorAgente.size()));//
            totalRow2.createCell(8).setCellValue(Math.round(tempoTotalRepeticoes / (double) dadosPorAgente.size()));//QTD - 4 posição antiga
            totalRow2.createCell(9).setCellValue(((totalPorcentoThpa / dadosPorAgente.size()) + "%").trim()); // Media THPA
            totalRow2.createCell(10).setCellValue(formatarSegundos(tempoTrabalhadoTotal / dadosPorAgente.size()).trim()); // Media TTL
            totalRow2.createCell(11).setCellValue(((totalPorcentoTtl / dadosPorAgente.size()) + "%").trim()); // TTL
            totalRow2.createCell(12).setCellValue(formatarSegundos(tempoTotalTma / dadosPorAgente.size()).trim()); // TMA
            totalRow2.createCell(13).setCellValue(formatarSegundos(totalMtsl / dadosPorAgente.size()).trim()); // MTSL

            for (int i = 0; i <= 13; i++) {
                if (totalRow2.getCell(i) != null) {
                    totalRow2.getCell(i).setCellStyle(totalStyle); // coloca estilo azul em media  até 12
                }
            }

            //'sheet' planilha
            int totalLinhas = sheet.getLastRowNum(); // Última linha preenchida

            // A média está na penúltima linha
            Row mediaRow = sheet.getRow(totalLinhas);           
            Cell mediaCell = mediaRow.getCell(8); // média QTD 7
            //System.out.println("Nuemro celula 7 "+mediaCell);
            String mediaValor = mediaCell.toString().trim(); // Pega o valor da média como string

            Cell mediaCellTtl = mediaRow.getCell(10); // média TTL posição aintga 9
            String mediaTtl = mediaCellTtl.toString().trim(); // Pega o valor da média como string

            // Percorre todas as linhas da coluna 4 (exceto somatório e média)
            for (int i = 1; i < totalLinhas - 1; i++) { // Começa em 1 para ignorar o cabeçalho
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(8); // Pega a célula da coluna 8 QTD
                    if (cell != null) {
                        String valorCelula = cell.toString().trim(); // Converte o valor para string
                        boolean isTime = false; // Defina `true` se for tempo

                        // Obtém o estilo condicional
                        CellStyle style = EstilosExcel.getConditionalColumnStyle(workbook, valorCelula, mediaValor,
                                isTime);
                        cell.setCellStyle(style); // Aplica a formatação na célula
                    }

                    Cell tempo = row.getCell(10); // pega valor da celula na coluna 10 TTL.
                    //System.out.println("TTL " + tempo);
                    if (tempo != null) {
                        // String valorCelulaTtl = tempo.toString().trim();
                        String valorCelulaTtl = tempo.getCellType() == CellType.STRING
                                ? tempo.getStringCellValue().trim()
                                : EstilosExcel.TIME_FORMAT.format(tempo.getDateCellValue());
                        mediaTtl = mediaCellTtl.getCellType() == CellType.STRING
                                ? mediaCellTtl.getStringCellValue().trim()
                                : EstilosExcel.TIME_FORMAT.format(mediaCellTtl.getDateCellValue());
                        boolean isTime = true;

                        // Obtém o estilo condicional
                        CellStyle style = EstilosExcel.getConditionalColumnStyle(workbook, valorCelulaTtl, mediaTtl,
                                isTime);
                        tempo.setCellStyle(style); // Aplica a formatação na célula

                    }                  

                }
                
            }

            // Salvar o arquivo Excel
            try (FileOutputStream fileOut = new FileOutputStream("operadoras.xlsx")) {
                sheet.autoSizeColumn(15);
                workbook.write(fileOut);
            }

            System.out.println("Arquivo Excel 'operadoras.xlsx' gerado com sucesso!");

        } catch (IOException e) {
            e.printStackTrace();  
            
            try {
                Files.write(Paths.get("erro.log"), Arrays.asList(e.toString()), StandardOpenOption.APPEND);
            } catch (IOException ex) {
                ex.printStackTrace(); // Se der erro ao escrever no arquivo, exibe no console
            }           
        }
    }

    private static long calcularDiferencaSegundos(String inicio, String fim) {
        try {
            if (inicio == null || fim == null || inicio.isEmpty() || fim.isEmpty()) {
                return 0;
            }
            long segundosInicio = converterParaSegundos(inicio.split(" ")[1]);
            long segundosFim = converterParaSegundos(fim.split(" ")[1]);
            return Math.max(0, segundosFim - segundosInicio);
        } catch (Exception e) {
            System.err.println("Erro ao calcular diferença: " + e.getMessage());
            return 0;
        }
    }

    private static long converterParaSegundos(String hora) {
        try {
            String[] partes = hora.split(":");
            return Integer.parseInt(partes[0]) * 3600 + Integer.parseInt(partes[1]) * 60 + Integer.parseInt(partes[2]);
        } catch (Exception e) {
            return 0;
        }
    }
    public static long durationToSeconds(String duration) {
        String[] units = duration.split(":");
        long hours = Long.parseLong(units[0]);
        long minutes = Long.parseLong(units[1]);
        long seconds = Long.parseLong(units[2]);
        return hours * 3600 + minutes * 60 + seconds;
    }

    private static String formatarSegundos(long segundos) {
        long horas = (long) segundos / 3600;
        long minutos = (long) (segundos % 3600) / 60;
        long seg = (long) segundos % 60;
        return String.format("%02d:%02d:%02d", horas, minutos, seg);
    }

    private static int calcularPorcentagem(long tempoTrabalhadoSegundos, long tempoTotalSegundos) {

        return (int) Math.round(((double) tempoTrabalhadoSegundos / tempoTotalSegundos) * 100);
    }

    // Método para realizar a divisão simples
    public static long calcularDivisao(long valor1, long valor2) {
        if (valor2 == 0) {
            throw new ArithmeticException("Divisão por zero não é permitida.");
        }
        return valor1 / valor2;
    }

    // Método para aplicar o estilo de borda a todas as células de uma linha
    public static void applyBorderStyleToRow(Row row, CellStyle style) {
        for (Cell cell : row) {
            cell.setCellStyle(style);
        }
    }

     // Função para extrair a parte numérica do agente (exemplo: "ope1002" → 1002)
     private static int extrairNumeroAgente(String agente) {
        return Integer.parseInt(agente.replaceAll("\\D+", ""));
    }

}
