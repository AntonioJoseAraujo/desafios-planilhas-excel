# Projeto Planilha Vagas Python
### Monte um programa que cria a planilha com a estrutura abaixo e permita que o usuário adicione dados a essa planilha
#(todos dados devem ser recebidos no formato de string msm)
#* Empresa
#* Vaga
#* Data da Aplicação
#* Retorno Extra
### Nome da planilha "Acompanhamento de Vagas"
### Nome do sheet - vagas
### Mantenha apenas o sheet que foi criado
import openpyxl

planilha = openpyxl.Workbook()
del planilha['Sheet']
planilha.create_sheet('vagas')
pagina_vagas = planilha['vagas']
pagina_vagas.append(['Empresa', 'Vaga', 'Data da Aplicação', 'Retorno'])

cont = 's'
while cont == 's':
    empresa = input('Empresa: ')
    vaga = input('Nome da Vaga: ')
    data = input('Data da Aplicação: ')
    retorno = input('Retorno: ')
    pagina_vagas.append([empresa,vaga,data,retorno])
    cont= input('Adicionar nova vaga? (s/n): ')

planilha.save('Acompanhamento de Vagas.xlsx')