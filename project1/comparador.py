import pandas as pd

arquivo1 = 'arquivo1.xlsx'
arquivo2 = 'arquivo2.xlsx'
relatorio = 'relatorio_diferencas.txt'

# Leitura dos arquivos
df1 = pd.read_excel(arquivo1)
df2 = pd.read_excel(arquivo2)

# Selecionar as 3 primeiras colunas (por posição)
df1_sub = df1.iloc[:, 0:3].fillna('')
df2_sub = df2.iloc[:, 0:3].fillna('')

# Pega o nome real das colunas para mostrar no relatório
colunas = df1_sub.columns.tolist()

# Ajustar tamanho para igualar linhas
max_linhas = max(len(df1_sub), len(df2_sub))
df1_sub = df1_sub.reindex(range(max_linhas)).fillna('')
df2_sub = df2_sub.reindex(range(max_linhas)).fillna('')

with open(relatorio, 'w', encoding='utf-8') as f:
    f.write('📄 RELATÓRIO DE DIFERENÇAS ENTRE ARQUIVOS EXCEL (3 primeiras colunas)\n')
    f.write('='*60 + '\n')

    diferencas_encontradas = False

    for i in range(max_linhas):
        for idx, coluna in enumerate(colunas):
            valor1 = str(df1_sub.iat[i, idx]).strip()
            valor2 = str(df2_sub.iat[i, idx]).strip()

            if valor1 != valor2:
                diferencas_encontradas = True
                f.write(f"\n🔍 Linha {i+1}, Coluna '{coluna}' (posição {idx+1}):\n")
                f.write(f" - Valor em {arquivo1}: {valor1}\n")
                f.write(f" - Valor em {arquivo2}: {valor2}\n")

    if not diferencas_encontradas:
        f.write("✅ Nenhuma diferença encontrada entre os arquivos nas 3 primeiras colunas.\n")

print(f"✅ Comparação concluída. Relatório salvo em: {relatorio}")
