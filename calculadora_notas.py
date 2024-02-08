import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Define o escopo e as credenciais para a API do Google Sheets
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials_path = 'C:\\Users\\David\\Documents\\Engenharia de Software teste\\engenharia-desoftware-davidjez-7e0e51fd1379.json'  # Caminho para o seu arquivo JSON de credenciais
creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
client = gspread.authorize(creds)

# Abre a planilha e seleciona a primeira planilha
spreadsheet = client.open('Cópia de Engenharia de Software - Desafio: David Jezrrel Ribeiro Neves')
worksheet = spreadsheet.sheet1

def calcular_nota(aluno):
    total_aulas = 60
    total_faltas_permitidas = total_aulas * 0.25

    nome_aluno = aluno[1]
    faltas = int(aluno[2])
    
    # Verifica se as notas são valores numéricos antes de converter
    notas = []
    for nota in aluno[3:6]:
        try:
            nota_int = int(nota)
            notas.append(nota_int)
        except ValueError:
            return "Erro: As notas devem ser valores numéricos", 0

    # Calcula a média das notas
    media_notas = sum(notas) / len(notas)

    # Verifica se o aluno foi reprovado por faltas
    if faltas > total_faltas_permitidas:
        return "Reprovado por Falta", 0

    # Calcula a situação com base na média
    if media_notas < 5:
        return "Reprovado por Nota", 0
    elif 5 <= media_notas < 7:
        # Calcula a nota necessária para aprovação final
        naf = max(0, (10 - media_notas) * 2)
        return "Exame Final", int(naf)
    else:
        return "Aprovado", 10  # Retorna 10 como a nota final para alunos aprovados


# Processa cada linha de aluno
for i, linha in enumerate(worksheet.get_all_values()):
    if i == 0 or not linha[0].isdigit():  # Ignora a linha de cabeçalho e outras linhas que não são dados de alunos
        continue
    
    id_aluno = int(linha[0])
    nome_aluno = linha[1]
    situacao, nota_final = calcular_nota(linha)
    
    # Escreve a situação e a nota final na planilha
    worksheet.update_cell(i + 1, 6, situacao)
    worksheet.update_cell(i + 1, 7, nota_final)
    
    print(f"{nome_aluno} teve sua situação e nota final atualizadas: {situacao}, {nota_final}")

print("Situações e notas finais de todos os alunos atualizadas com sucesso.")
