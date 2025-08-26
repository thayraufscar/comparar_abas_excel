import pandas as pd
import re
import unicodedata

def normalize_text(text):
    if not text or not isinstance(text, str):
        return ""
    # Remove acentos
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    # Minúsculas
    text = text.lower()
    # Substitui qualquer caractere não alfanumérico (pontuação, hífen, barra etc.) por espaço
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    # Remove espaços extras
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# === CONFIGURAÇÃO ===
arquivo_entrada = "entrada.xlsx"
aba1 = "Sheet1"   # primeira aba
aba2 = "Sheet2"   # segunda aba
coluna = "Titulo" # coluna a comparar
arquivo_saida = "comparacao_abas.xlsx"

# === PROCESSO ===
# Lê as duas abas
df1 = pd.read_excel(arquivo_entrada, sheet_name=aba1)
df2 = pd.read_excel(arquivo_entrada, sheet_name=aba2)

# Normaliza os textos apenas para comparar
df1["__normalizado__"] = df1[coluna].apply(normalize_text)
df2["__normalizado__"] = df2[coluna].apply(normalize_text)

# Conjuntos para comparar
set1 = set(df1["__normalizado__"])
set2 = set(df2["__normalizado__"])

# Diferenças
somente1 = df1[df1["__normalizado__"].isin(set1 - set2)].drop(columns="__normalizado__")
somente2 = df2[df2["__normalizado__"].isin(set2 - set1)].drop(columns="__normalizado__")

# Interseção (registros em comum)
comum = df1[df1["__normalizado__"].isin(set1 & set2)].drop(columns="__normalizado__")

# Exporta resultado em três abas
with pd.ExcelWriter(arquivo_saida) as writer:
    somente1.to_excel(writer, sheet_name="Somente_" + aba1, index=False)
    somente2.to_excel(writer, sheet_name="Somente_" + aba2, index=False)
    comum.to_excel(writer, sheet_name="Comum", index=False)

print(f"✅ Comparação concluída. Resultado em {arquivo_saida}")
