import javax.swing.*;
import java.io.*;
import java.util.Map;

public class ProcessarArquivoCSV {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(JanelaPrincipal::new);
       /* JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(true); // Permite selecionar múltiplos arquivos
        fileChooser.setDialogTitle("Selecione os arquivos CSV");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Arquivos CSV", "csv"));

        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File[] arquivosSelecionados = fileChooser.getSelectedFiles();

            if (arquivosSelecionados.length != 2) {
                System.out.println("Por favor, selecione exatamente dois arquivos.");
                return;
            }

            File arquivo3601 = null;
            File arquivo3502 = null;

            // Identifica qual arquivo é qual
            for (File arquivo : arquivosSelecionados) {
                String nomeArquivo = arquivo.getName();
                if (nomeArquivo.contains("3601")) {
                    arquivo3601 = arquivo;
                } else if (nomeArquivo.contains("3502")) {
                    arquivo3502 = arquivo;
                } else {
                    System.out.println("Arquivo desconhecido: " + nomeArquivo);
                    return;
                }
            }

            if (arquivo3601 == null || arquivo3502 == null) {
                System.out.println("Erro: Certifique-se de selecionar um arquivo 3601 e um arquivo 3502.");
                return;
            }

            // Processar o arquivo 3601
            ProcessarArquivo3601 processar3601 = new ProcessarArquivo3601();
            Map<String, String[]> dados3601 = processar3601.processarArquivo(arquivo3601.getAbsolutePath());

            // Processar o arquivo 3502 e integrar os dados do 3601
            tratarArquivo3502(arquivo3502, dados3601);

            System.out.println("Todos os arquivos foram processados com sucesso. Saída gerada em operadoras.xlsx.");
        } else {
            System.out.println("Nenhum arquivo selecionado.");
        }
    }

    // Método para processar o arquivo 3502 diretamente
    private static void tratarArquivo3502(File arquivo3502, Map<String, String[]> dados3601) {
        try {
            ProcessarArquivo3502 tratarArquivo3502 = new ProcessarArquivo3502();
            tratarArquivo3502.processarDados(arquivo3502.getAbsolutePath(), dados3601);
        } catch (Exception e) {
            System.out.println("Erro ao processar o arquivo " + arquivo3502.getName() + ": " + e.getMessage());           
        }
            */
    }
}