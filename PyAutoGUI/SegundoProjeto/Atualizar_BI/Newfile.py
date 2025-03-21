from graphviz import Digraph

# Create a new directed graph
dot = Digraph(comment='Processo Tecnologia da Informação', format='pdf')

# Add nodes for each task and event
dot.node('StartEvent_Recebimento_Chamados', 'Recebimento de Chamados pelo GLPI', shape='ellipse')
dot.node('Task_Classificacao_Chamados', 'Classificação dos Chamados', shape='box')
dot.node('Task_Atribuicao_Chamados', 'Atribuição dos Chamados', shape='box')

# SubProcess Tratamento dos Chamados
dot.node('SubProcess_Tratamento_Chamados', 'Tratamento dos Chamados', shape='box')
dot.node('Task_Site', 'Site (Manutenção e atualizações)', shape='box', style='filled', color='lightgrey')
dot.node('Task_Polibras', 'Polibras (Gestão e resolução de problemas)', shape='box', style='filled', color='lightgrey')
dot.node('Task_ERP_SAP', 'ERP SAP (Suporte e manutenção)', shape='box', style='filled', color='lightgrey')
dot.node('Task_Apitolo', 'Apitolo (Gestão e manutenção da rede de arquivos)', shape='box', style='filled', color='lightgrey')
dot.node('Task_Hardware', 'Hardware (Gestão de equipamentos e manutenção)', shape='box', style='filled', color='lightgrey')
dot.node('Task_Suporte_Lojas', 'Suporte às Lojas (Atendimento a chamados)', shape='box', style='filled', color='lightgrey')
dot.node('Task_BI', "BI's (Criação e manutenção)", shape='box', style='filled', color='lightgrey')

dot.node('Task_Resolucao_Fechamento', 'Resolução e Fechamento dos Chamados', shape='box')

# Reuniões e Compras
dot.node('StartEvent_Necessidade_Equipamentos', 'Necessidade de novos equipamentos', shape='ellipse')
dot.node('Task_Reuniao_Discussao_Aprovacao', 'Reunião para Discussão e Aprovação da Compra', shape='box')
dot.node('Task_Compra_Equipamentos', 'Compra de Equipamentos', shape='box')
dot.node('Task_Configuracao_Distribuicao', 'Configuração e Distribuição de Equipamentos', shape='box')

# Revisão e Melhoria Contínua
dot.node('Task_Revisao_Processos', 'Revisão Periódica dos Processos', shape='box')
dot.node('EndEvent_Melhoria_Continua', 'Melhoria Contínua', shape='ellipse')

# Add edges with labels
dot.edges(['StartEvent_Recebimento_Chamados', 'Task_Classificacao_Chamados'])
dot.edge('Task_Classificacao_Chamados', 'Task_Atribuicao_Chamados')
dot.edge('Task_Atribuicao_Chamados', 'SubProcess_Tratamento_Chamados')
dot.edge('SubProcess_Tratamento_Chamados', 'Task_Resolucao_Fechamento')

dot.edge('StartEvent_Necessidade_Equipamentos', 'Task_Reuniao_Discussao_Aprovacao')
dot.edge('Task_Reuniao_Discussao_Aprovacao', 'Task_Compra_Equipamentos')
dot.edge('Task_Compra_Equipamentos', 'Task_Configuracao_Distribuicao')

dot.edge('Task_Resolucao_Fechamento', 'Task_Revisao_Processos')
dot.edge('Task_Revisao_Processos', 'EndEvent_Melhoria_Continua')

# Render the graph to a PDF file
file_path_pdf = '/mnt/data/Processo_Technologia_da_Informacao.pdf'
dot.render(file_path_pdf, view=False)

file_path_pdf
