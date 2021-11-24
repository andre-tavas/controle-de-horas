# controle-de-horas

## Descrição
Sistema envolvendo as ferramentas Toggl e Google Sheets para controle de horas e análise de atividades de organizações. Foi criado inicialmente para o projeto de extensão Inderios Consultoria. Fornece informações relacionadas à quantidade de tempo gasta por cada um em diferentes categorias (como projetos ou equipes das quais a pessoa participa), além do tempo gasto por todas as pessoas em cada equipe ou projeto por mês ou no total.
Utiliza o Toggl para registro dos dados de atividades e o Google Sheets para armazenamento, processamento e visualização destes dados. Os dados são importados para a planilha diariamente por meio de um script do Google acionado por um acionador. Esse script utiliza a API do toggl.

## Apresentação da aplicação
[![Miniatura para vídeo do youtube](http://img.youtube.com/vi/m1BXgaKSFe4/0.jpg)](https://youtu.be/m1BXgaKSFe4)


### Pontos de melhoria
- Limpeza do código
- Funcionalidade de alerta imediato no email do gestor e do usuário que registra uma atividade muito longa no Toggl (é comum usuários esquecerem o cronômetro ligado após o término de uma atividade
- Escalabilidade da ferramenta para outras organizações
- Aplicação de ferramentas de mineração de processos para análises mais aprofundadas (pm4py, BupaR, etc)
