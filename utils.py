import pandas as pd
from openpyxl import Workbook
import os

def create_template():
    filename = "template_notas.xlsx"
    if os.path.exists(filename):
        return filename
    
    # Criando dados de exemplo
    data = [
        {"Turma": "9º Ano A", "Nome do Aluno": "João Silva", "Unidade Curricular": "Matemática", "Nota": 8.5, "Frequência Geral (%)": 95},
        {"Turma": "9º Ano A", "Nome do Aluno": "João Silva", "Unidade Curricular": "Português", "Nota": 7.0, "Frequência Geral (%)": 95},
        {"Turma": "9º Ano A", "Nome do Aluno": "Maria Oliveira", "Unidade Curricular": "Matemática", "Nota": 9.5, "Frequência Geral (%)": 98},
        {"Turma": "9º Ano A", "Nome do Aluno": "Maria Oliveira", "Unidade Curricular": "Português", "Nota": 9.0, "Frequência Geral (%)": 98},
        {"Turma": "9º Ano A", "Nome do Aluno": "Pedro Santos", "Unidade Curricular": "Matemática", "Nota": 6.5, "Frequência Geral (%)": 85},
        {"Turma": "9º Ano A", "Nome do Aluno": "Pedro Santos", "Unidade Curricular": "Português", "Nota": 6.0, "Frequência Geral (%)": 85},
    ]
    
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Template criado: {filename}")
    return filename

if __name__ == "__main__":
    create_template()
