import java.util.Date;

class Agente {
    private String nome;
    private Date menorHoraInicial;
    private Date maiorHoraFinal;
    private int repeticoes;
    private long tempoTotal;

    public Agente(String nome) {
        this.nome = nome;
        this.menorHoraInicial = new Date(Long.MAX_VALUE);
        this.maiorHoraFinal = new Date(Long.MIN_VALUE);
        this.repeticoes = 0;
        this.tempoTotal = 0;
    }

    public String getNome() { return nome; }
    public Date getMenorHoraInicial() { return menorHoraInicial; }
    public Date getMaiorHoraFinal() { return maiorHoraFinal; }
    public int getRepeticoes() { return repeticoes; }
    public long getTempoTotal() { return tempoTotal; }

    public void atualizarDados(Date horaInicial, Date horaFinal, long tempoTrabalhado) {
        if (horaInicial.before(menorHoraInicial)) menorHoraInicial = horaInicial;
        if (horaFinal.after(maiorHoraFinal)) maiorHoraFinal = horaFinal;
        repeticoes++;
        tempoTotal += tempoTrabalhado;
    }
}