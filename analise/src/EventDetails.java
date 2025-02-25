public class EventDetails {
    private String loginInicial;
    private String loginFinal;
    private long duracao;  // Campo para armazenar a duração
    private int quantidade; // Campo para armazenar a quantidade

    // Getters e Setters
    public String getLoginInicial() {
        return loginInicial;
    }

    public void setLoginInicial(String loginInicial) {
        this.loginInicial = loginInicial;
    }

    public String getLoginFinal() {
        return loginFinal;
    }

    public void setLoginFinal(String loginFinal) {
        this.loginFinal = loginFinal;
    }

    public long getDuracao() {
        return duracao;
    }

    public void setDuracao(long duracao) {
        this.duracao = duracao;
    }

    public int getQuantidade() {
        return quantidade;
    }

    public void setQuantidade(int quantidade) {
        this.quantidade = quantidade;
    }

    // Métodos para incrementar a quantidade e a duração
    public void incrementarQuantidade() {
        this.quantidade++;
    }

    public void incrementarDuracao(long duracaoSegundos) {
        this.duracao += duracaoSegundos;
    }
}
