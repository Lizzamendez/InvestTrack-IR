# InvestTrack-IR
📊 Agregador de Dados para Imposto de Renda no Excel

Este projeto foi desenvolvido como parte de um desafio prático na DIO (Digital Innovation One). O objetivo é criar uma ferramenta robusta no Microsoft Excel para centralizar, organizar e validar dados financeiros essenciais para a Declaração de Imposto de Renda de Pessoa Física (DIRPF).



🚀 Funcionalidades

O projeto transforma uma planilha comum em uma ferramenta automatizada com:



Configurações Centrais (Config): Banco de dados com códigos COMPE de instituições financeiras brasileiras e categorias de ativos.

Gestão de Lançamentos: Registro de operações de Compra e Venda com cálculo automático de custos operacionais.

Controle de Proventos: Registro detalhado de Dividendos, JCP e Rendimentos, permitindo separar fluxos isentos de tributáveis.

Consolidado Automático: Resumo em tempo real de ativos, calculando o Preço Médio Ponderado e a quantidade atual em custódia.

UX com VBA: Sistema de navegação por botões dinâmicos que utilizam macros para melhorar a experiência do usuário.

🛠️ Tecnologias Utilizadas

Microsoft Excel: Motor principal da ferramenta.

VBA (Visual Basic for Applications): Para automação de interface e movimentação de objetos.

Fórmulas Avançadas: Uso de SOMASES, PROCX (ou PROCV) e lógica condicional para tratamento de erros.

Markdown: Para documentação técnica no GitHub.

📐 Estrutura do Projeto

1. Inteligência de Cálculos

O coração da ferramenta é o cálculo do Preço Médio, essencial para o IR. A fórmula utilizada no consolidado garante que o custo de aquisição seja calculado corretamente:



Excel



=SE(Qtd_Atual>0; Total_Investido / Total_Qtd_Comprada; 0)

2. Automação de Interface (VBA)

A planilha conta com um menu interativo. O código VBA abaixo é responsável por mover o marcador visual e alternar entre as abas:



VBA



Sub NavegarPara(aba As String, posicaoX As Double)

    Dim shp As Shape

    Set shp = ActiveSheet.Shapes("MarcadorMenu")

    shp.Left = posicaoX

    Sheets(aba).Activate

End Sub

📋 Como utilizar

Config: Verifique se os bancos e tipos de ativos estão cadastrados.

Lançamentos: Insira suas notas de corretagem (Data, Ativo, Operação, Qtd e Preço).

Proventos: Registre os valores recebidos conforme seus informes de rendimentos.

Consolidado: Acompanhe seu preço médio e posição atual de forma automática.

✍️ Autor

Desenvolvido por Lizza Mendez durante a Formação na DIO.
