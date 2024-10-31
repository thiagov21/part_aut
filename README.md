Descrição do Código para Particionamento de Planilhas Excel
Este script em Python é projetado para particionar uma planilha Excel em várias planilhas separadas com base na coluna "Filial". O processo utiliza as bibliotecas pandas e openpyxl para ler, manipular e formatar os dados de maneira eficiente.

Funcionalidades Principais:
Leitura da Planilha Base: O código inicia carregando uma planilha Excel existente, que contém informações sobre diferentes filiais, com colunas como 'Filial', 'Nome', 'Data' e 'Valor'.

Criação de Diretórios: Antes de criar os arquivos Excel separados, o código verifica se a pasta de destino existe. Se não existir, ele cria essa pasta automaticamente.

Filtragem de Dados: Para cada filial única identificada na coluna 'Filial', o script filtra os dados correspondentes e cria um novo arquivo Excel contendo apenas os dados daquela filial específica.

Formatação da Planilha: Após gerar cada arquivo Excel, o código aplica formatação:

Centralização do Conteúdo: O conteúdo de todas as células é centralizado para melhorar a estética da planilha.
Ajuste Automático das Larguras das Colunas: O script calcula a largura necessária para cada coluna com base no comprimento do conteúdo, garantindo que os dados sejam apresentados de forma legível.
Formatação em Tabela: Cada planilha é formatada como uma tabela, o que proporciona uma visualização mais organizada e facilita a leitura dos dados.
Mensagens de Conclusão: Ao final do processo, uma mensagem é exibida no console para indicar que o particionamento e a formatação das planilhas foram concluídos com sucesso.

Uso
Para utilizar este script, certifique-se de ter as bibliotecas pandas e openpyxl instaladas em seu ambiente Python. Atualize os caminhos de entrada e saída conforme necessário para refletir sua estrutura de arquivos.

Este código é especialmente útil em cenários onde grandes volumes de dados precisam ser organizados e apresentados de forma clara, permitindo que usuários ou equipes de diferentes filiais acessem facilmente suas informações específicas.
