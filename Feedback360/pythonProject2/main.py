# PRÓMIXO PASSO, ENVIAR EMAIL

# Graficos
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# EMAIL
import mimetypes
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

qntPerguntas = 11;

df = pd.read_excel("FEEDBACK360_2.xlsx")
df2 = df.iloc[:, 3:]  # eliminando as colunas de data e id (3 primeiras colunas)

[avaliacoes, nColunas] = df2.shape  # numero de linhas (avaliações) e colunas (quesitos)

# Filtrando campos gerados pelo forms
df2.drop(df2.filter(regex='^(Pontos|Comentários).+$'), axis=1, inplace=True)
df2.to_excel("saida1.xlsx", index=False)

# Obtendo lista de equipes
equipes = df2['Escolha o seu grupo:'].unique()

# Obtendo lista de alunos
alunos = df2['Nome'].unique()

print(alunos)
print(equipes)

# Obtendo apenas o conceito para cada avaliação (sem o texto)
df2 = df2.drop(columns=['Nome', 'Total de pontos', 'Digite seu nome completo:', 'Escolha o seu grupo:']) # Temporario
for column in df2.columns[1:nColunas - 1]:
    if(column != 'Escreva o e-mail do integrante que será avaliado:' and not(column.startswith("Escreva o e-mail do integrante que será avaliado:"))):
        df2[column] = df2[column].astype(str).str[0]

# Convertendo conceitos em notas numéricas
replacement_mapping_dict = {"A": 5, "B": 4, "C": 3, "D": 2, "E": 1}
df2.replace(replacement_mapping_dict, inplace=True)

# Ajustando as categorias
df2.columns = df2.columns.str[:1]

df2.to_excel("saida2.xlsx", index=False)

# Notas do aluno para ele mesmo
df3 = df2.iloc[: , 1:qntPerguntas+1]
df3 = df3.groupby(level=0, axis=1, sort=False).median()
print(df3)

df3.rename(columns={'1': 'Solução'}, inplace=True)
df3.rename(columns={'2': 'Métodos'}, inplace=True)
df3.rename(columns={'3': 'Protótipo'}, inplace=True)
df3.rename(columns={'4': 'Organização'}, inplace=True)
df3.rename(columns={'5': 'Apresentação'}, inplace=True)
df3.rename(columns={'6': 'Competência Pessoal'}, inplace=True)
df3.rename(columns={'7': 'Normas'}, inplace=True)
df3.rename(columns={'8': 'Pesquisa'}, inplace=True)

df3.to_excel("saida3.xlsx", index=False)

# Notas da equipe para cada aluno
df4 = df2.iloc[:,qntPerguntas+1:]

i=0
while i < 4:
    df_aux1 = df4.iloc[:, :(qntPerguntas+1)*1]
    df_aux2 = df4.iloc[:, 12:24]
    df_aux3 = df4.iloc[:, 24:36]
    df_aux4 = df4.iloc[:, 36:48]
    i+=1


pdList = [df_aux1,df_aux2,df_aux3,df_aux4]
new_df = pd.concat(pdList)

df4.to_excel("saida4.xlsx", index=False)

# Removendo emails vazios
new_df.replace("", inplace=True)
new_df.dropna(subset = ["E"], inplace=True)
new_df = new_df.groupby(['E'], as_index=False).median()
alunos2 = new_df['E'].unique()
new_df = new_df.iloc[:, 1:]
new_df = new_df.groupby(level=0, axis=1, sort=False).median()

new_df.rename(columns={'1': 'Solução'}, inplace=True)
new_df.rename(columns={'2': 'Métodos'}, inplace=True)
new_df.rename(columns={'3': 'Protótipo'}, inplace=True)
new_df.rename(columns={'4': 'Organização'}, inplace=True)
new_df.rename(columns={'5': 'Apresentação'}, inplace=True)
new_df.rename(columns={'6': 'Competência Pessoal'}, inplace=True)
new_df.rename(columns={'7': 'Normas'}, inplace=True)
new_df.rename(columns={'8': 'Pesquisa'}, inplace=True)

new_df.to_excel("saida5.xlsx", index=False)


# ======================================
# ENVIAR EMAIL
# ======================================
def enviar_email(to_email, subject, number):
    de = 'gabrielss2406@hotmail.com'

    message = MIMEMultipart()
    message['From'] = de
    message['To'] = to_email
    message['Subject'] = subject
    message.attach(MIMEText("Relatório: " + alun, 'html'))

    cam_arquivo = alun + ".pdf"
    attchment = open(cam_arquivo, 'rb')

    att = MIMEBase('application', 'octet-stream')
    att.set_payload(attchment.read())
    encoders.encode_base64(att)

    att.add_header('Content-Disposition', f'attachment; filename=Relatorio.pdf')
    attchment.close()
    message.attach(att)

    smtp = smtplib.SMTP('smtp.office365.com', 587)
    smtp.starttls()
    smtp.login('gabrielss2406@hotmail.com', 'Inatel#2022')
    smtp.sendmail(message['From'], message['To'], message.as_string())
    smtp.quit()


# =====================================
# GRÁFICOS
# ======================================

# Salvando os graficos
def save_multi_image(filename, aluno):
    pp = PdfPages(filename)
    fig_nums = plt.get_fignums()
    figs = [plt.figure(n) for n in fig_nums]
    figs[0 + (aluno*2)].savefig(pp, format='pdf')
    #figs[1 + (aluno*2)].savefig(pp, format='pdf')
    pp.close()


# Guarda os valores em ordem alfabetica
values1 = df3.values.tolist() # Valores auto avaliação
values2 = new_df.values.tolist() # Valores avaliação dos colegas

# Escolhendo aluno
aluno = -1

# Listando categorias
categories = list(new_df)
categories = [el.replace('\xa0',' ') for el in categories] # Erro com \xa0


# Auto Avaliação
for alun in alunos:
    aluno += 1
    N = 5

    # Criando lista dos valores (proprio - 1 / colegas - 2
    values1[aluno] += values1[aluno][:1]

    # Gerando os angulos do grafico
    angles = [n / float(N) * 2 * 3.14 for n in range(N)]
    angles += angles[:1]

    # Gráfico 1
    # Inicializando os graficos
    fig = plt.figure()
    ax1 = fig.add_subplot(111, polar=True)

    # Desenhando linhas dos angulos e adicionando label
    plt.xticks(angles[:-1], categories, color='grey', size=8)
    ax1.set_title('Avaliação própria - ' + alun)

    # Desenhando linhas circulares
    ax1.set_rlabel_position(0)
    plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="red", size=8)
    plt.ylim(0, 5)

    # Plot do grafico
    ax1.plot(angles, values1[aluno], linewidth=1, linestyle='solid')

    # Preenchendo o grafico
    ax1.fill(angles, values1[aluno], 'b', alpha=0.1)
    # Mostrando os graficos
    fig1 = plt.figure()

    filename = alun + ".pdf"
    save_multi_image(filename, aluno)
    #enviar_email("gabrielss2406@hotmail.com", "Teste", alun)


# Limpando plots
for plot in plt.get_fignums():
    plt.close(plot)

# Escolhendo aluno
aluno = -1

# Alunos avaliados por colegas
for alun in alunos2:
    print(alun)
    aluno += 1
    N = 5

    # Criando lista dos valores (proprio - 1 / colegas - 2
    values2[aluno] += values2[aluno][:1]

    # Gerando os angulos do grafico
    angles = [n / float(N) * 2 * 3.14 for n in range(N)]
    angles += angles[:1]

    # Gráfico 1
    # Inicializando os graficos
    fig_ = plt.figure()
    ax2 = fig_.add_subplot(111, polar=True)

    # Desenhando linhas dos angulos e adicionando label
    plt.xticks(angles[:-1], categories, color='grey', size=8)
    ax2.set_title('Avaliação colegas - ' + alun)

    # Desenhando linhas circulares
    ax2.set_rlabel_position(0)
    plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="red", size=8)
    plt.ylim(0, 5)

    # Plot do grafico
    ax2.plot(angles, values2[aluno], linewidth=1, linestyle='solid')

    # Preenchendo o grafico
    ax2.fill(angles, values2[aluno], 'b', alpha=0.1)
    # Mostrando os graficos
    fig2 = plt.figure()

    filename = alun + "-COLEGAS.pdf"
    save_multi_image(filename, aluno)
    #enviar_email("gabrielss2406@hotmail.com", "Teste", alun)

