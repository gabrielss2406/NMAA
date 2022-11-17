# PRÓMIXO PASSO, ENVIAR EMAIL
import functions
# Graficos
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

disciplina = "G021"
qntPerguntas = 9
alunosAtencao = []
graphys = functions.generate_graphs_quantity(qntPerguntas)

functions.create_paste(disciplina)

df = pd.read_excel("./planilhas/Avaliação 360 - G021(1-4).xlsx")
df2 = df.iloc[:, 3:]  # eliminando as colunas de data e id (3 primeiras colunas)

[avaliacoes, nColunas] = df2.shape  # numero de linhas (avaliações) e colunas (quesitos)

# Filtrando campos gerados pelo forms
df2.drop(df2.filter(regex='^(Pontos|Comentários).+$'), axis=1, inplace=True)

# Obtendo lista de equipes
equipes = df2['Selecione o grupo que será avaliado:'].unique()

# Obtendo apenas o conceito para cada avaliação (sem o texto)
df2 = df2.drop(columns=['Email', 'Nome', 'Total de pontos']) # Temporario
for column in df2.columns[1:nColunas - 1]:
    if(column != 'Selecione o grupo que será avaliado:'):
        df2[column] = df2[column].astype(str).str[0]

# Convertendo conceitos em notas numéricas
replacement_mapping_dict = {"A": 5, "B": 4, "C": 3, "D": 2, "E": 1}
df2.replace(replacement_mapping_dict, inplace=True)

# Ajustando as categorias
df2.columns = df2.columns.str[:3]
df2.to_excel("saida.xlsx")

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


# Guarda os valores
df2 = df2.drop(columns=["Sel"])
values1 = df2.values.tolist() # Valores auto avaliação
values3 = [float(sum(col))/len(col) for col in zip(*values1)]

# Listando categorias
categories = list(df2)
categories = [el.replace('\xa0',' ') for el in categories] # Erro com \xa0

# Analisando dados
functions.analyse_all(values3,categories,disciplina)
functions.analyse_groups(values1,equipes,disciplina)
functions.analyse_notes_2(values1,categories,equipes,disciplina)

aluno = -1
# Auto Avaliação
for alun in equipes:
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
        ax1.set_title('Avaliação - '+ str(parte) + " - " + alun)

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



# ========================================
# GERAR GRAFICOS TOTAL
# ========================================
isAluno=False
for equipe in equipes:
    functions.concat_plots(equipe,graphys,isAluno,disciplina)

# ========================================
# GERAR GRAFICOS MEDIA
# ========================================
# Limpando plots
for plot in plt.get_fignums():
    plt.close(plot)

categories_cpy = categories.copy()
eNes = 0
aluno = 1
parte=0
print(values3)
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
    ax1.set_title('Media da turma - ' + str(parte) + " - " + alun)

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
txts = ['Pontos fracos e fortes da turma', 'Grupos_Melhor_Pior', 'Grupos_NotasBaixas', 'Grupos_NotasAltas']
functions.concat_txts(txts,disciplina)