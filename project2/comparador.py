import pandas as pd

def comparar_excel(arquivo1, arquivo2):
    relatorio_linhas = []
    try:
        df1 = pd.read_excel(arquivo1)
        df2 = pd.read_excel(arquivo2)
    except Exception as e:
        print(f"❌ Erro ao abrir os arquivos: {e}")
        return None

    df1_sub = df1.iloc[:, 0:3].fillna('')
    df2_sub = df2.iloc[:, 0:3].fillna('')
    colunas = df1_sub.columns.tolist()

    max_linhas = max(len(df1_sub), len(df2_sub))
    df1_sub = df1_sub.reindex(range(max_linhas)).fillna('')
    df2_sub = df2_sub.reindex(range(max_linhas)).fillna('')

    diferencas_encontradas = False

    relatorio_linhas.append('📄 RELATÓRIO DE DIFERENÇAS ENTRE ARQUIVOS EXCEL (3 primeiras colunas)')
    relatorio_linhas.append('='*60)

    for i in range(max_linhas):
        for idx, coluna in enumerate(colunas):
            valor1 = str(df1_sub.iat[i, idx]).strip()
            valor2 = str(df2_sub.iat[i, idx]).strip()

            if valor1 != valor2:
                diferencas_encontradas = True
                relatorio_linhas.append(f"\n🔍 Linha {i+1}, Coluna '{coluna}' (posição {idx+1}):")
                relatorio_linhas.append(f" - Valor em {arquivo1}: {valor1}")
                relatorio_linhas.append(f" - Valor em {arquivo2}: {valor2}")

    if not diferencas_encontradas:
        relatorio_linhas.append("✅ Nenhuma diferença encontrada entre os arquivos nas 3 primeiras colunas.")

    return '\n'.join(relatorio_linhas)


def main():
    print("=== Comparador de arquivos Excel ===")
    arquivo1 = input("Digite o caminho do primeiro arquivo Excel: ").strip()
    arquivo2 = input("Digite o caminho do segundo arquivo Excel: ").strip()

    relatorio = comparar_excel(arquivo1, arquivo2)
    if relatorio is None:
        print("Erro na comparação. Verifique os arquivos e tente novamente.")
        return

    print("\n=== Relatório de Diferenças ===")
    print(relatorio)

    salvar = input("\nDeseja salvar este relatório em um arquivo TXT? (s/n): ").strip().lower()
    if salvar == 's':
        nome_arquivo = input("Digite o nome do arquivo TXT (ex: relatorio.txt): ").strip()
        try:
            with open(nome_arquivo, 'w', encoding='utf-8') as f:
                f.write(relatorio)
            print(f"✅ Relatório salvo com sucesso em '{nome_arquivo}'")
        except Exception as e:
            print(f"❌ Erro ao salvar o arquivo: {e}")
    else:
        print("Relatório não salvo.")

if __name__ == '__main__':
    main()
