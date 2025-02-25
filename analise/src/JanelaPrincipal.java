import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.net.URL;
import java.util.Map;

public class JanelaPrincipal extends JFrame {
    private JLabel lblResultado;
    private JProgressBar progressBar;

    public JanelaPrincipal() {
        setTitle("Processador de CSV");
        setSize(500, 350);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // Carrega e ajusta a imagem do logo
        URL imagemURL = getClass().getClassLoader().getResource("imagens/Bem-Vindo.png");
        if (imagemURL != null) {
            ImageIcon icone = new ImageIcon(imagemURL);
            Image imagem = icone.getImage().getScaledInstance(200, 100, Image.SCALE_SMOOTH);
            JLabel lblLogo = new JLabel(new ImageIcon(imagem), SwingConstants.CENTER);
            add(lblLogo, BorderLayout.NORTH);
        } else {
            System.err.println("Imagem não encontrada!");
        }

        // Criando painel de botões
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);

        JButton btnSelecionarArquivos = criarBotao("Selecionar Arquivos", "imagens/folder.png");
        JButton btnFinalizar = criarBotao("Finalizar", "imagens/exit.png");

        gbc.gridx = 0;
        gbc.gridy = 0;
        panel.add(btnSelecionarArquivos, gbc);

        gbc.gridx = 1;
        panel.add(btnFinalizar, gbc);
        
        add(panel, BorderLayout.CENTER);

        lblResultado = new JLabel("", SwingConstants.CENTER);
        add(lblResultado, BorderLayout.SOUTH);

        progressBar = new JProgressBar(0, 100);
        progressBar.setStringPainted(true);
        add(progressBar, BorderLayout.PAGE_END);

        // Eventos dos botões
        btnSelecionarArquivos.addActionListener(e -> selecionarArquivos());
        btnFinalizar.addActionListener(e -> finalizarAplicacao());

        setVisible(true);
    }

    private JButton criarBotao(String texto, String iconePath) {
        JButton botao = new JButton(texto);
        botao.setFont(new Font("Arial", Font.BOLD, 12));
        botao.setBackground(new Color(70, 130, 180));
        botao.setForeground(Color.WHITE);
        botao.setFocusPainted(false);
        botao.setBorder(BorderFactory.createRaisedBevelBorder());
        
        URL iconeURL = getClass().getClassLoader().getResource(iconePath);
        if (iconeURL != null) {
            botao.setIcon(new ImageIcon(new ImageIcon(iconeURL).getImage().getScaledInstance(20, 20, Image.SCALE_SMOOTH)));
        }
        return botao;
    }

    private void selecionarArquivos() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setFileFilter(new FileNameExtensionFilter("Arquivos CSV", "csv"));
        int resposta = fileChooser.showOpenDialog(this);

        if (resposta == JFileChooser.APPROVE_OPTION) {
            File[] arquivosSelecionados = fileChooser.getSelectedFiles();

            if (arquivosSelecionados.length != 2) {
                JOptionPane.showMessageDialog(this, "Por favor, selecione exatamente dois arquivos.", "Erro", JOptionPane.ERROR_MESSAGE);
                return;
            }

            new Thread(() -> {
                processarArquivos(arquivosSelecionados);
                JOptionPane.showMessageDialog(this, "Arquivo gerado com sucesso! Consulte sua pasta.", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
                SwingUtilities.invokeLater(() -> lblResultado.setText("Arquivos processados. Salvo em: " + obterCaminhoSaida()));
            }).start();
        }
    }

    private void processarArquivos(File[] arquivos) {
        int total = arquivos.length;
        Map<String, String[]> dados3601 = null;
    
        for (int i = 0; i < total; i++) {
            if (arquivos[i].getName().contains("3601")) {
                ProcessarArquivo3601 processar3601 = new ProcessarArquivo3601();
                dados3601 = processar3601.processarArquivo(arquivos[i].getAbsolutePath());
            }
        }
    
        for (int i = 0; i < total; i++) {
            if (arquivos[i].getName().contains("3502")) {
                if (dados3601 != null) {
                    tratarArquivo3502(arquivos[i], dados3601);
                } else {
                    JOptionPane.showMessageDialog(this, "Erro: O arquivo 3502 foi encontrado antes do 3601. Refaça a seleção.", "Erro", JOptionPane.ERROR_MESSAGE);
                    return;
                }
            }
    
            final int progressIndex = i; // Nova variável final
            try { Thread.sleep(1000); } catch (InterruptedException e) { e.printStackTrace(); }
            SwingUtilities.invokeLater(() -> progressBar.setValue((progressIndex + 1) * 100 / total)); // Use progressIndex aqui
        }
    }

    private String obterCaminhoSaida() {
        return System.getProperty("user.home") + "/Downloads/resultado.csv";
    }

    private static void tratarArquivo3502(File arquivo3502, Map<String, String[]> dados3601) {
        try {
            ProcessarArquivo3502 tratarArquivo3502 = new ProcessarArquivo3502();
            tratarArquivo3502.processarDados(arquivo3502.getAbsolutePath(), dados3601);
        } catch (Exception e) {
            System.err.println("Erro ao processar " + arquivo3502.getName() + ": " + e.getMessage());
        }
    }

    private void finalizarAplicacao() {
        int resposta = JOptionPane.showConfirmDialog(this, "Tem certeza que deseja finalizar?", "Confirmar saída", JOptionPane.YES_NO_OPTION);
        if (resposta == JOptionPane.YES_OPTION) {
            System.exit(0);
        }
    }
}
