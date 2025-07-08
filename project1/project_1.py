import pandas as pd

# Caminhos dos arquivos de entrada
arquivo1 = 'arquivo1.xlsx'
arquivo2 = 'arquivo2.xlsx'

# Caminho do arquivo de saída
relatorio = 'relatorio_diferencas.txt'

# Colunas a comparar
colunas_para_comparar = ['A', 'B', 'C']

# Leitura dos arquivos
df1 = pd.read_excel(arquivo1)
df2 = pd.read_excel(arquivo2)

# Verificação das colunas
for coluna in colunas_para_comparar:
    if coluna not in df1.columns or coluna not in df2.columns:
        print(f"❌ A coluna '{coluna}' não foi encontrada em um dos arquivos.")
        exit()

# Garantir que ambos tenham o mesmo número de linhas
max_linhas = max(len(df1), len(df2))
df1 = df1.reindex(range(max_linhas)).fillna('')
df2 = df2.reindex(range(max_linhas)).fillna('')

# Abrir o relatório para escrita
with open(relatorio, 'w', encoding='utf-8') as f:
    f.write('📄 RELATÓRIO DE DIFERENÇAS ENTRE ARQUIVOS EXCEL\n')
    f.write('='*60 + '\n')

    diferencas_encontradas = False

    for i in range(max_linhas):
        for coluna in colunas_para_comparar:
            valor1 = str(df1.at[i, coluna]).strip()
            valor2 = str(df2.at[i, coluna]).strip()

            if valor1 != valor2:
                diferencas_encontradas = True
                f.write(f"\n🔍 Linha {i+1}, Coluna '{coluna}':\n")
                f.write(f" - Valor em {arquivo1}: {valor1}\n")
                f.write(f" - Valor em {arquivo2}: {valor2}\n")

    if not diferencas_encontradas:
        f.write("✅ Nenhuma diferença encontrada entre os arquivos nas colunas A, B e C.\n")

print(f"✅ Comparação concluída. Relatório salvo em: {relatorio}")
