# Processador de Arquivo CSV para Relatório Excel

Este projeto em **Java** processa arquivos **CSV** contendo informações sobre chamadas telefônicas, aplicando regras específicas e extraindo informações sobre operadores e suas atividades, e gera um **relatório no formato Excel** (.xlsx) utilizando a biblioteca **Apache POI**.

## Funcionalidades

- Leitura de arquivos CSV contendo registros de chamadas de operadores.
- Processamento e extração de dados conforme regras específicas.
- Cálculo de tempos trabalhados, tempo total em ligação (TTL), tempo médio de atendimento (TMA) e MTSL.
- Agrupamento de dados por operador.
- Geração de uma planilha Excel com formatação, bordas e estilos condicionalmente aplicados.
- Exibição de totais e médias gerais na planilha.
- Aplicação de estilos personalizados às células do Excel.

## Estrutura do Projeto

O projeto é composto pelas seguintes classes principais:

- **`ProcessarArquivoCSV.java`**: Classe principal responsável pela execução do processamento dos arquivos CSV.
- **`Agente.java`**: Representa os agentes, armazenando informações relevantes extraídas dos arquivos CSV.
- **`EstilosExcel.java`**: Define os estilos utilizados na formatação do Excel gerado, aplicando cores, 
fontes e alinhamento.

## Tecnologias Utilizadas

- **Java SE 8+**
- **Apache POI** (para manipulação de arquivos Excel)
- **Swing JFileChooser** (para seleção do arquivo CSV)
- **Bibliotecas do Java Nativo**: `BufferedReader`, `SimpleDateFormat`, `HashMap`, `List`, entre outras.

## Como Executar

1. **Baixe ou clone o repositório**.
2. **Configure o Apache POI** no seu projeto.
3. **Compile e execute** a classe `ProcessarArquivoCSV`.
4. **Selecione um arquivo CSV** quando solicitado.
5. O sistema processará os dados e gerará um arquivo Excel com os resultados.

# Como Executar Terminal
1. Compile o projeto com:
   ```sh
   javac -cp .;lib/* src/*.java -d bin
   ```
2. Execute a classe principal:
   ```sh
   java -cp bin;lib/* ProcessarArquivoCSV

## Estrutura do Arquivo CSV
O arquivo lido deve ter a seguinte estrutura:

```
Data de Atendimento;Data Inicial da Chamada;Agente;Tempo de Conversação;Tempo Total;Tempo de Pós-atendimento;DDD;Origem;Destino;Data Final da Chamada
```

## Estrutura Cabeçalho de Saída do Arquivo CSV

O arquivo CSV deve conter as seguintes colunas:

1. **DATA** - Data da chamada
2. **OPERADORA** - Nome do operador
3. **1ºL** - Horário inicial
4. **TULD** - Horário final
5. **QTD** - Quantidade de chamadas
6. **%THPA** - Percentual de horas trabalhadas
7. **TTL** - Tempo total de ligação
8. **%TTL** - Percentual do tempo total em ligação
9. **TMA** - Tempo médio de atendimento
10. **MTSL** - Maior tempo de sessão de ligação

## Estrutura do Arquivo Excel Gerado

- **Planilha nomeada como "Análise das OPE - DD-MM-YYYY"**.
- **Cabeçalho formatado** com estilos personalizados.
- **Cada linha representa um operador** e seus tempos calculados.
- **Linha de totais e médias** ao final da planilha.

## Dependências Necessárias

Adicione as bibliotecas do Apache POI ao seu projeto. No caso de uso com **Maven**, inclua no `pom.xml`:

```xml
<dependencies>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>4.1.2</version>
    </dependency>
</dependencies>
```

## Melhorias Futuras

- Interface Gráfica para seleção de parâmetros.
- Suporte a mais formatos de entrada (ex: JSON, XML).
- Geração de gráficos no Excel.

## Autor

Desenvolvido por Helvisklei em 26/02/2025.


