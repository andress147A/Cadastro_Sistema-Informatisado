Funcionamento 

Option Explicit: Obriga a declaração explícita de todas as variáveis, ajudando a evitar erros de digitação e referências a variáveis não declaradas.

Botões de Comando (CommandButtonX_Click):

CommandButton1 e CommandButton2: Abrem os formulários UserForm2 e UserForm3, respectivamente.
CommandButton3 e CommandButton4: Exportam os dados das planilhas "Planilha1" e "Planilha2" para arquivos JSON.
CommandButton5 e CommandButton6: Exportam os dados das mesmas planilhas para XML.
CommandButton7 e CommandButton8: Geram gráficos baseados nos dados das planilhas.
ExportarParaJson(nomePlanilha As String):

Converte os dados da planilha especificada em formato JSON.
Salva o JSON em um arquivo no mesmo diretório da planilha.
ExportarParaXML(nomePlanilha As String):

Converte os dados para XML e salva como um arquivo.
GerarGrafico(nomePlanilha As String):

Analisa os dados e cria um gráfico de frequência dos trechos (partida-chegada).
EncontrarTrechosMaisMenosRequisitados(nomePlanilha As String):

Identifica os trechos mais e menos populares da planilha.
O código automatiza a exportação de dados e a geração de gráficos a partir de informações contidas nas planilhas do Excel.
