"""Baixa a planilha do Google Drive e grava em dados/dados.xlsx se o conteúdo mudou."""
import io
import os
import pathlib
import urllib.request
import zipfile

DRIVE_ID = os.environ.get("DRIVE_ID", "1sgly0VR5hVdCLex1q16zoAhBSfYTLQqx")
SAIDA = pathlib.Path("dados/dados.xlsx")


def conteudo(dados: bytes) -> dict:
    # docProps guarda datas de salvamento; fora dele está o que importa para o site.
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        return {n: z.read(n) for n in z.namelist() if not n.startswith("docProps/")}


url = f"https://drive.google.com/uc?export=download&id={DRIVE_ID}"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
novo = urllib.request.urlopen(req, timeout=60).read()

if not novo.startswith(b"PK"):
    raise SystemExit("O Drive não devolveu uma planilha. Confira se o compartilhamento está como 'qualquer pessoa com o link'.")
novo_conteudo = conteudo(novo)
if "xl/workbook.xml" not in novo_conteudo:
    raise SystemExit("O arquivo baixado não é uma planilha .xlsx válida.")

if SAIDA.exists() and conteudo(SAIDA.read_bytes()) == novo_conteudo:
    print("Planilha sem mudanças.")
else:
    SAIDA.write_bytes(novo)
    print(f"Planilha atualizada ({len(novo)} bytes).")
