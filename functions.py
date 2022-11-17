import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PyPDF2 import PdfMerger
import os

def create_paste(disciplina):
    if not os.path.isdir("./" + disciplina):
        os.makedirs("./"+disciplina)
        os.mkdir("./"+disciplina+"/Analises")
        os.mkdir("./"+disciplina+"/Graficos")
        os.mkdir("./"+disciplina+"/Graficos/Alunos")
        os.mkdir("./"+disciplina+"/Graficos/Outros")


def generate_graphs_quantity(questions):
    graficos = []
    while (questions != 0):
        if (questions % 6 == 0):
            graficos.append(6)
            questions -= 6
        elif (questions % 5 == 0):
            graficos.append(5)
            questions -= 5
        elif (questions % 4 == 0):
            graficos.append(4)
            questions -= 4
        elif (questions % 3 == 0):
            graficos.append(3)
            questions -= 3
        elif (questions == 7):
            graficos.append(4)
            questions -= 4
        else:
            graficos.append(6)
            questions -= 6
    return graficos


def send_email(to_email, subject, number, alun):
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


def concat_plots(alun, N, isAluno, disciplina):
    merger = PdfMerger()
    i=0
    for n in N:
        i+=1
        merger.append(alun+"-"+str(i)+".pdf")
    if(isAluno):
        i=0
        for n in N:
            i+=1
            merger.append(alun+""+str(i)+"-C.pdf")
        merger.write("./"+disciplina+"/Graficos/Alunos/"+alun + ".pdf")
    else:
        merger.write("./"+disciplina+"/Graficos/Outros/"+alun + ".pdf")
    merger.close()

    i=0
    for n in N:
        i+=1
        os.remove(alun+"-"+str(i)+".pdf")
    if(isAluno):
        i = 0
        for n in N:
            i+=1
            os.remove(alun+""+str(i)+"-C.pdf")


def analyse_notes(values1,values2,categories,alunos,disciplina): # values1 (autoavaliacao) - values2 (colegas)
    notasBaixas = []
    desviosAltos = []
    notasAltas = []
    i=0
    #print(alunos)
    for notas in values1: # cada notas é de um aluno
        j=0
        #print(notas)
        for cat in notas:
            if cat < 2:
                notasBaixas.append(alunos[i]+" - "+categories[j]+" - AutoAvaliação - Nota muito baixa (< 2)\n")
            elif cat>4.5:
                notasAltas.append(alunos[i]+" - "+categories[j]+" - AutoAvaliação - Nota muito alta (> 4.5)\n")
            if values2[i][j] < 2:
                notasBaixas.append(alunos[i]+" - "+categories[j]+" - Avaliação Colegas - Nota muito baixa (< 2)\n")
            elif values2[i][j]>4.5:
                notasAltas.append(alunos[i]+" - "+categories[j]+" - Avaliação Colegas - Nota muito alta (> 4.5)\n")

            if (cat - values2[i][j]) > 2:
                desviosAltos.append(alunos[i]+" - "+categories[j]+" - Grande Desvio (> 2) - AutoAvaliação maior que a dos colegas\n")
            elif (cat - values2[i][j]) < -2:
                desviosAltos.append(alunos[i]+" - "+categories[j]+" - Grande Desvio (> 2) - Avaliação dos colegas maior que AutoAvaliação\n")

            j+=1
        i+=1

    k = 0
    arquivo = open("./"+disciplina+"/Analises/Alunos_NotasBaixas.txt", 'w')
    for a in notasBaixas:
        arquivo.write(notasBaixas[k])
        k += 1

    l = 0
    arquivo = open("./"+disciplina+"/Analises/Alunos_GrandesDesvios.txt", 'w')
    arquivo.write('\n')
    for a in desviosAltos:
        arquivo.write(desviosAltos[l])
        l += 1

    p = 0
    arquivo = open("./"+disciplina+"/Analises/Alunos_NotasAltas.txt", 'w')
    for a in notasAltas:
        arquivo.write(notasAltas[p])
        p += 1


def analyse_all(values,categories,disciplina):
    print(values)
    pontosFortesFracos = []
    maiores = sorted(range(len(values)), key=lambda i: values[i], reverse=True)
    menores = sorted(range(len(values)), key=lambda i: values[i], reverse=False)
    i=0
    pontosFortesFracos.append("Pontos fortes da turma: \n")
    while(i<3):
        pontosFortesFracos.append(categories[maiores[i]]+" - Nota Media: "+str(values[maiores[i]])+"\n")
        i+=1

    pontosFortesFracos.append("\nPontos fracos da turma: \n")
    i=0
    while(i<3):
        pontosFortesFracos.append(categories[menores[i]]+" - Nota Media: "+str(values[menores[i]])+"\n")
        i+=1

    k = 0
    arquivo = open("./"+disciplina+"/Analises/Pontos fracos e fortes da turma.txt", 'w')
    for a in pontosFortesFracos:
        arquivo.write(pontosFortesFracos[k])
        k += 1


def analyse_groups(values, groups,disciplina):
    i=0
    media=0
    notas = []
    for group in groups:
        media=0
        j=0
        for nota in values[i]:
            media+=nota
            j+=1
        media /= j
        notas.append(media)
        i+=1
    k = 0

    arquivo = open("./"+disciplina+"/Analises/Grupos_Melhor_Pior.txt", 'w')
    arquivo.write("Grupo com maior nota média: " +groups[values.index(max(values))]+" - %.2f" % notas[values.index(max(values))])
    arquivo.write("\nGrupo com menor nota média: " +groups[values.index(min(values))]+" - %.2f" % notas[values.index(min(values))])


def analyse_notes_2(values1,categories,alunos,disciplina): # values1 (autoavaliacao) - values2 (colegas)
    notasBaixas = []
    notasAltas = []
    notasBaixas.append('\n')
    i=0
    for notas in values1: # cada notas é de um aluno
        j=0
        #print(notas)
        for cat in notas:
            if cat < 2:
                notasBaixas.append(alunos[i]+" - "+categories[j]+" - Nota muito baixa (< 2)\n")
            elif cat>4.5:
                notasAltas.append(alunos[i]+" - "+categories[j]+" - Nota muito alta (> 4.5)\n")
            j+=1
        i+=1

    k = 0
    arquivo = open("./"+disciplina+"/Analises/Grupos_NotasBaixas.txt", 'w')
    for a in notasBaixas:
        arquivo.write(notasBaixas[k])
        k += 1

    p = 0
    arquivo = open("./"+disciplina+"/Analises/Grupos_NotasAltas.txt", 'w')
    for a in notasAltas:
        arquivo.write(notasAltas[p])
        p += 1


def concat_txts(txts, disciplina):
    with open('./'+disciplina+"/Analises/Relatorio.txt", "w") as file:
        for temp in txts:
            with open('./'+disciplina+'/Analises/'+temp+'.txt', "r") as t:
                file.writelines(t)
                file.write('\n')
        for temp in txts:
            os.remove('./' + disciplina + '/Analises/' + temp + '.txt')

    #doc = aw.Document('./'+disciplina+'/Analises/geral.txt')
    #doc.save('./'+disciplina+'/Analises/Relatório.pdf')
