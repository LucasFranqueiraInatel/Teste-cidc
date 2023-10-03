# Data Analysis and Reporting Tool
Este é um script Python que lida com a leitura, manipulação e geração de relatórios a partir de planilhas do Excel usando as bibliotecas openpyxl e pandas.

## Pré-requisitos
Certifique-se de ter as seguintes bibliotecas instaladas:

- openpyxl: Necessária para gerar e manipular planilhas do Excel.

* pandas: Necessária para manipular os DataFrames.
Você pode instalá-las usando pip:

## Funcionalidades
### Leitura de Dados:

Função **read_data()** para ler dados de uma planilha Excel.
### Geração de Relatórios:

Função **generate_report()** para gerar um relatório combinando informações de duas planilhas Excel.
### Análises de Dados:

**Função quality_0()** para identificar sites com qualidade zero.

**Função mbps_more_80()** para identificar sites com velocidade maior que 80 Mbps.

**Função mbps_less_10()** para identificar sites com velocidade menor que 10 Mbps.
### Verificação de Sites Ausentes:

Função **sites_not_in_results()** para identificar sites presentes em uma planilha, mas ausentes em outra.

### Observação: 
A ideia central foi deixar o código o mais legível possível, minimizando a necessidade de muitos comentários. A estrutura das funções foi pensada para facilitar a compreensão e adaptação a diferentes planilhas. Houve uma certa dúvida sobre a origem dos dados no início, se seriam do novo excel gerado ou do anterior, levando à decisão de organizar o código em funções para fácil implementação e chamada.
Um exemplo disso é na função **alert_yes()** como na planilha result.xlsx não existe a coluna Alerts foi desenvolvida a função bastando chama-la.

Segue a questão a qual fiquei com duvida sobre qual planilha era a correta.
![image](https://github.com/JoaoPedroAM/Teste-cidc/assets/64596663/26032f48-cf82-4aad-8913-91d5cfd685de)
