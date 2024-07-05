# Análise de Métricas de Email

Esta aplicação em Python permite analisar métricas de campanhas de email a partir de arquivos CSV ou Excel. O programa calcula diversas métricas de eficácia, como taxa de abertura, taxa de cliques, taxa de descadastros e feedback score, e gera um relatório de desempenho. O resultado é salvo em uma nova planilha Excel com colunas adicionais e larguras ajustadas para melhor legibilidade.

## Funcionalidades

- Leitura de arquivos CSV e Excel
- Cálculo de métricas de email: taxa de abertura, taxa de cliques, taxa de descadastros e feedback score
- Geração de relatório de desempenho para cada linha de dados
- Ajuste automático da largura das colunas na planilha de saída
- Registro de logs de erros e atividades

## Instalação

1. Clone o repositório:
    ```sh
    git clone https://github.com/seu-usuario/analise-metricas-email.git
    cd analise-metricas-email
    ```

2. Crie e ative um ambiente virtual (opcional, mas recomendado):
    ```sh
    python -m venv venv
    source venv/bin/activate   # Para sistemas Unix
    venv\Scripts\activate      # Para Windows
    ```

3. Instale as dependências:
    ```sh
    pip install -r requirements.txt
    ```

## Uso

1. Execute a aplicação:
    ```sh
    python analista.py
    ```

2. Na interface gráfica, clique no botão "Selecionar Arquivo" para escolher o arquivo CSV ou Excel contendo os dados das campanhas de email.

3. A aplicação processará o arquivo, calculará as métricas e salvará uma nova planilha Excel com as colunas adicionais e ajustadas.

4. O novo arquivo será salvo no mesmo diretório do arquivo original, com "_com_metricas" adicionado ao nome do arquivo.

## Estrutura do Projeto

- `analista.py`: Script principal da aplicação.
- `requirements.txt`: Arquivo de dependências Python.
- `email_metrics.log`: Arquivo de log gerado pela aplicação.

## Dependências

- pandas
- openpyxl
- tkinter

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues e pull requests.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
