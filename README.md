# controle-de-horas

## Descrição
Sistema envolvendo as ferramentas Toggl e Google Sheets para controle de horas e análise de atividades de organizações. Foi criado inicialmente para um projeto de extensão focado na prestação de consultoria para pequenos negócios. Fornece informações relacionadas à quantidade de tempo gasta por cada um em diferentes categorias (como projetos ou equipes das quais a pessoa participa), além do tempo gasto por todas as pessoas em cada equipe ou projeto por mês ou no total.
Utiliza o Toggl para registro dos dados de atividades e o Google Sheets para armazenamento, processamento e visualização destes dados. Os dados são importados para a planilha diariamente por meio de um script do Google acionado por um acionador. Esse script utiliza a API do toggl.

Foi desenvolvido um modelo com instruções e possibilidade de configurar algumas funcionalidades de forma personalizada para escalabilidade da ferramenta. Este modelo encontra-se [neste link](https://docs.google.com/spreadsheets/d/1kE5G6uHdg8cpJoO2tLHtZNpv5K77PECEwbAWK-EREKk/edit#gid=2001691719), caso queira utilizar a ferramenta, sinta-se a vontade para realizar uma cópia do arquivo.

O código que comunica com a API do toggl, processa os dados e registra em uma aba foram adaptados de: https://stargazerllc.medium.com/programatically-import-toggl-time-data-to-google-sheets-invoice-449530b5f2d5

## Videos
### Apresentação da aplicação
[![Miniatura para vídeo do youtube](http://img.youtube.com/vi/aaTwLi2cZuA/0.jpg)](https://youtu.be/aaTwLi2cZuA)
### Como ligar a automação
[![Miniatura para vídeo do youtube](http://img.youtube.com/vi/nJ8gcaR7SzA/0.jpg)](https://youtu.be/nJ8gcaR7SzA)
