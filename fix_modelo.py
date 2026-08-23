import shutil, openpyxl

ORIG = "Simulador_FII modelo padrão.xlsx"
DEST = "Simulador_FII.xlsx"

# 1) Usa o modelo padrão como base de formatação
shutil.copyfile(ORIG, DEST)

wb = openpyxl.load_workbook(DEST)
perfis = wb["Perfis"]

# 2) Corrige a coluna CHAVE (A) que estava vazia -> causa do #N/D no PROCV
for r in range(3, 21):
    perfis[f"A{r}"] = f"=B{r}&\"-\"&C{r}"

# 3) Ajusta o texto de dica (referenciava H21, mas o perfil fica em L41)
ws = wb["Simulador"]
if ws["K52"].value and "H21" in str(ws["K52"].value):
    ws["K52"] = ("Dica: altere os valores em azul para simular cenários e use o menu "
                 "suspenso em 'Perfil de risco' (L41) para recalcular a alocação por tipo de FII.")

wb.save(DEST)
print("Base de formatação aplicada (modelo padrão) e CHAVE corrigida.")
