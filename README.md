🤖 SAP Data Automation – Extração, Conversão e Envio via API
💡 Visão Geral

Este projeto realiza a automação completa do ciclo de dados no SAP, desde a extração de diversas tabelas até o envio de informações consolidadas para uma API externa.

Desenvolvido em Python com integração SAP GUI Scripting, o processo executa consultas automáticas, exporta relatórios em Excel, converte, trata e estrutura os dados em formato JSON, enviando tudo via requisição HTTP.

📊 Tabelas SAP Automatizadas
Tabela	Descrição	Finalidade
EKKO	Cabeçalho de pedidos de compra	Base de documentos
EKKO (Contrato)	Pedidos vinculados a contrato	Controle de contrato
LFA1	Dados do fornecedor	Identificação e nome
EKPO	Itens do pedido de compra	Detalhes linha a linha
MARA	Dados gerais do material	Tipo, grupo e descrição
ADR6	Endereços de e-mail dos usuários	Comunicação
EKET	Datas de entrega	Cronograma de fornecimento
USR21	Vínculo entre usuários SAP e e-mails	Relacionamento interno
MM03	Visualização do material	Detalhamento adicional
ME23N	Visualização de pedido	Validação cruzada
⚙️ Funcionalidades

✅ Conexão automática ao SAP GUI via win32com.client
✅ Exportação programada de relatórios .XLS
✅ Conversão automática para .XLSX
✅ Manipulação e limpeza com Pandas
✅ Criação de dicionários Python → JSON
✅ Envio de JSON via requisição HTTP (API GET - POST - PUT)
✅ Logs detalhados de execução e tratamento de erros
✅ Estrutura modular (cada tabela tem seu script próprio)


📈 Benefícios

Reduz tempo de extração manual no SAP
Padroniza e automatiza consultas complexas
Integração direta com sistemas externos (Coupa, APIs REST, etc.)
Gera histórico e rastreabilidade por logs

📝 Licença

Este projeto é de uso interno e educacional, não distribuível publicamente sem autorização da autora.
