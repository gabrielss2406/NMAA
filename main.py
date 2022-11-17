# PRÓMIXO PASSO, ENVIAR EMAIL
import functions
# Graficos
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

disciplina = "G007"
qntPerguntas = 5
alunosAtencao = []
graphys = functions.generate_graphs_quantity(qntPerguntas)

functions.create_paste(disciplina)

df = pd.read_excel("./planilhas/FEEDBACK360_4.xlsx")
df2 = df.iloc[:, 3:]  # eliminando as colunas de data e id (3 primeiras colunas)

[avaliacoes, nColunas] = df2.shape  # numero de linhas (avaliações) e colunas (quesitos)

# Filtrando campos gerados pelo forms
df2.drop(df2.filter(regex='^(Pontos|Comentários).+$'), axis=1, inplace=True)

# Obtendo lista de equipes
equipes = df2['Escolha o seu grupo:'].unique()

# Obtendo lista de alunos
alunos = df2['Email'].unique()

# Obtendo apenas o conceito para cada avaliação (sem o texto) [df5 - grupos, df2 - alunos]
df5 = df2.drop(columns=['Email', 'Nome', 'Total de pontos', 'Digite seu nome completo:', 'Escreva o e-mail do integrante que será avaliado:']) # Temporario
for column in df5.columns[1:nColunas - 1]:
    if(column.startswith('Escreva o e-mail do integrante que será avaliado:')):
        df5 = df5.drop(columns=column)
    else:
        df5[column] = df5[column].astype(str).str[0]

df2 = df2.drop(columns=['Nome', 'Total de pontos', 'Digite seu nome completo:', 'Escolha o seu grupo:']) # Temporario
for column in df2.columns[1:nColunas - 1]:
    if(column != 'Escreva o e-mail do integrante que será avaliado:' and not(column.startswith("Escreva o e-mail do integrante que será avaliado:"))):
        df2[column] = df2[column].astype(str).str[0]

# Convertendo conceitos em notas numéricas
replacement_mapping_dict = {"A": 5, "B": 4, "C": 3, "D": 2, "E": 1}
df2.replace(replacement_mapping_dict, inplace=True)
df5.replace(replacement_mapping_dict, inplace=True)

# Ajustando as categorias
df2.columns = df2.columns.str[:3]
df5.columns = df5.columns.str[:3]

# Notas do aluno para ele mesmo
df3 = df2.iloc[: , 1:qntPerguntas+1]
df3 = df3.groupby(level=0, axis=1, sort=False).mean()

# Notas da equipe para cada aluno
df4 = df2.iloc[:,qntPerguntas+1:]

#df2.to_excel("saida1.xlsx", index=False) # alunos
#df5.to_excel("./"+disciplina+"/saida1.xlsx", index=False) # grupos
#df3.to_excel("saida4.xlsx", index=False) # alunos - autoavaliacao
#df4.to_excel("saida5.xlsx", index=False) # alunos

# Limpando campos vazios
i=0
while i < 4:
    df_aux1 = df4.iloc[:, :(qntPerguntas+1)*1]
    df_aux2 = df4.iloc[:, (qntPerguntas+1)*1:(qntPerguntas+1)*2]
    df_aux3 = df4.iloc[:, (qntPerguntas+1)*2:(qntPerguntas+1)*3]
    df_aux4 = df4.iloc[:, (qntPerguntas+1)*3:(qntPerguntas+1)*4]
    i+=1


pdList = [df_aux1,df_aux2,df_aux3,df_aux4]
new_df = pd.concat(pdList)
# Removendo emails vazios
new_df.replace("", inplace=True)
new_df.dropna(subset = ["Esc"], inplace=True)
new_df = new_df.groupby(['Esc'], as_index=False).mean()
alunos2 = new_df['Esc'].unique()
new_df = new_df.iloc[:, 1:]
new_df = new_df.groupby(level=0, axis=1, sort=False).mean()

#new_df.to_excel("saida3.xlsx", index=False)

# =====================================
# GRÁFICOS
# ======================================

# Salvando os graficos
def save_multi_image(filename, aluno):
    pp = PdfPages(filename)
    fig_nums = plt.get_fignums()
    figs = [plt.figure(n) for n in fig_nums]
    figs[0].savefig(pp, format='pdf')
    pp.close()


# Guarda os valores em ordem alfabetica
df3.to_excel("saidamkgbjr.xlsx")
values1 = df3.values.tolist() # Valores auto avaliação
values2 = new_df.values.tolist() # Valores avaliação dos colegas
values3 = values1+values2
avg = [float(sum(col))/len(col) for col in zip(*values3)]
values3 = avg

# Listando categorias
categories = list(new_df)
categories = [el.replace('\xa0',' ') for el in categories] # Erro com \xa0

# Analisando dados
functions.analyse_all(values3,categories,disciplina)
functions.analyse_notes(values1,values2,categories,alunos,disciplina)

aluno = -1
# Auto Avaliação
for alun in alunos:
    categories_cpy = categories.copy()
    eNes=0
    aluno += 1
    parte = 0
    for N in graphys:
        parte+=1
        # Escolhendo aluno
        values1_aux = values1[aluno].copy()
        values1_aux = values1_aux[:N]
        values1_aux += values1[aluno][:1+eNes]


        # Gerando os angulos do grafico
        angles = [n / float(N) * 2 * 3.14 for n in range(N)]
        angles += angles[:1]

        # Gráfico 1
        # Inicializando os graficos
        fig = plt.figure()
        ax1 = fig.add_subplot(111, polar=True)

        categoriesNow = categories_cpy[:N]
        del categories_cpy[0:N]

        # Desenhando linhas dos angulos e adicionando label
        plt.xticks(angles[:-1], categoriesNow, color='grey', size=8)
        ax1.set_title('Avaliação própria - '+ str(parte) + " - " + alun)

        # Desenhando linhas circulares
        ax1.set_rlabel_position(0)
        plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="red", size=8)
        plt.ylim(0, 5)

        # Plot do grafico
        ax1.plot(angles, values1_aux[:N+1], linewidth=1, linestyle='solid')

        # Preenchendo o grafico
        ax1.fill(angles, values1_aux[:N+1], 'b', alpha=0.1)
        # Mostrando os graficos
        fig1 = plt.figure()

        filename = alun + "-" + str(parte) +".pdf"
        save_multi_image(filename, aluno)
        eNes += N
        del values1[aluno][0:N]
        for plot in plt.get_fignums():
            plt.close(plot)

    #enviar_email("gabrielss2406@hotmail.com", "Teste", alun)


aluno = -1
# Alunos avaliados por colegas
for alun in alunos2:
    categories_cpy = categories.copy()
    eNes = 0
    aluno += 1
    parte=0
    for N in graphys:
        parte+=1
        # Escolhendo aluno
        values2_aux = values2[aluno].copy()
        values2_aux = values2_aux[:N]
        values2_aux += values2[aluno][:1 + eNes]

        # Gerando os angulos do grafico
        angles = [n / float(N) * 2 * 3.14 for n in range(N)]
        angles += angles[:1]

        # Gráfico 1
        # Inicializando os graficos
        fig = plt.figure()
        ax1 = fig.add_subplot(111, polar=True)

        categoriesNow = categories_cpy[:N]
        del categories_cpy[0:N]

        # Desenhando linhas dos angulos e adicionando label
        plt.xticks(angles[:-1], categoriesNow, color='grey', size=8)
        ax1.set_title('Avaliação dos colegas - '+ str(parte) + " - " + alun)

        # Desenhando linhas circulares
        ax1.set_rlabel_position(0)
        plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="red", size=8)
        plt.ylim(0, 5)

        # Plot do grafico
        ax1.plot(angles, values2_aux[:N + 1], linewidth=1, linestyle='solid')

        # Preenchendo o grafico
        ax1.fill(angles, values2_aux[:N + 1], 'b', alpha=0.1)
        # Mostrando os graficos
        fig1 = plt.figure()

        filename = alun + "" + str(parte) + "-C.pdf"
        save_multi_image(filename, aluno)
        eNes += N
        del values2[aluno][0:N]
        for plot in plt.get_fignums():
            plt.close(plot)

    # enviar_email("gabrielss2406@hotmail.com", "Teste", alun)


# ========================================
# GERAR GRAFICOS TOTAL
# ========================================
isAluno=True
for alun in alunos:
    functions.concat_plots(alun,graphys,isAluno,disciplina)

# ========================================
# GERAR GRUPOS
# ========================================
# Limpando plots
for plot in plt.get_fignums():
    plt.close(plot)

df5 = df5.groupby('Esc').mean()
df5 = df5.groupby(level=0, axis=1, sort=True).mean()
aluno = -1
aux=0
values4 = df5.values.tolist() # Valores grupos

for equip in equipes:
    equipes[aux] = equip[:7]
    aux+=1

functions.analyse_groups(values4,equipes,disciplina)
# Analisando grupos
aluno = -1
for alun in equipes:
    categories_cpy = categories.copy()
    eNes = 0
    aluno += 1
    parte=0
    for N in graphys:
        parte+=1
        # Escolhendo aluno
        values4_aux = values4[aluno].copy()
        values4_aux = values4_aux[:N]
        values4_aux += values4[aluno][:1 + eNes]

        # Gerando os angulos do grafico
        angles = [n / float(N) * 2 * 3.14 for n in range(N)]
        angles += angles[:1]

        # Gráfico 1
        # Inicializando os graficos
        fig = plt.figure()
        ax1 = fig.add_subplot(111, polar=True)

        categoriesNow = categories_cpy[:N]
        del categories_cpy[0:N]

        # Desenhando linhas dos angulos e adicionando label
        plt.xticks(angles[:-1], categoriesNow, color='grey', size=8)
        ax1.set_title('Media - ' + str(parte) + " - " + alun)

        # Desenhando linhas circulares
        ax1.set_rlabel_position(0)
        plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="red", size=8)
        plt.ylim(0, 5)

        # Plot do grafico
        ax1.plot(angles, values4_aux[:N + 1], linewidth=1, linestyle='solid')

        # Preenchendo o grafico
        ax1.fill(angles, values4_aux[:N + 1], 'b', alpha=0.1)
        # Mostrando os graficos
        fig1 = plt.figure()

        filename = alun + "-" + str(parte) + ".pdf"
        save_multi_image(filename, aluno)
        eNes += N
        del values4[aluno][0:N]
        for plot in plt.get_fignums():
            plt.close(plot)
    #enviar_email("gabrielss2406@hotmail.com", "Teste", alun)
isAluno=False
for equip in equipes:
    functions.concat_plots(equip,graphys,isAluno,disciplina)

# GERAR MEDIA GERAL

categories_cpy = categories.copy()
eNes = 0
aluno = 1
parte=0
for N in graphys:
    parte+=1
    # Escolhendo aluno
    values3_aux = values3.copy()
    values3_aux = values3_aux[:N]
    values3_aux += values3[:1 + eNes]

    # Gerando os angulos do grafico
    angles = [n / float(N) * 2 * 3.14 for n in range(N)]
    angles += angles[:1]

    # Gráfico 1
    # Inicializando os graficos
    fig = plt.figure()
    ax1 = fig.add_subplot(111, polar=True)

    categoriesNow = categories_cpy[:N]
    del categories_cpy[0:N]

    # Desenhando linhas dos angulos e adicionando label
    plt.xticks(angles[:-1], categoriesNow, color='grey', size=8)
    ax1.set_title('Media da turma - ' + str(parte))

    # Desenhando linhas circulares
    ax1.set_rlabel_position(0)
    plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="red", size=8)
    plt.ylim(0, 5)

    # Plot do grafico
    ax1.plot(angles, values3_aux[:N + 1], linewidth=1, linestyle='solid')

    # Preenchendo o grafico
    ax1.fill(angles, values3_aux[:N + 1], 'b', alpha=0.1)
    # Mostrando os graficos
    fig1 = plt.figure()

    filename = "Media Turma -" + str(parte) + ".pdf"
    save_multi_image(filename, aluno)
    eNes += N
    del values3[0:N]
    for plot in plt.get_fignums():
            plt.close(plot)

functions.concat_plots("Media Turma ",graphys,False,disciplina)
txts = ['Pontos fracos e fortes da turma', 'Grupos_Melhor_Pior', 'Alunos_GrandesDesvios' ,'Alunos_NotasBaixas', 'Alunos_NotasAltas']
functions.concat_txts(txts,disciplina)