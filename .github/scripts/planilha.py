"""Baixa a planilha de cada série (dados/series.json) do Google Drive e grava em
dados/series/<id>/dados.xlsx quando o conteúdo mudou."""
import io
import json
import pathlib
import urllib.request
import zipfile

SERIES = json.loads(pathlib.Path("dados/series.json").read_text("utf-8"))["series"]


def conteudo(dados: bytes) -> dict:
    # docProps guarda datas de salvamento; fora dele está o que importa para o site.
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        return {n: z.read(n) for n in z.namelist() if not n.startswith("docProps/")}


def baixar(serie):
    saida = pathlib.Path(f"dados/series/{serie['id']}/dados.xlsx")
    url = f"https://drive.google.com/uc?export=download&id={serie['drive']}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    novo = urllib.request.urlopen(req, timeout=60).read()
    if not novo.startswith(b"PK"):
        raise RuntimeError("o Drive não devolveu uma planilha; confira se o compartilhamento está como 'qualquer pessoa com o link'")
    novo_conteudo = conteudo(novo)
    if "xl/workbook.xml" not in novo_conteudo:
        raise RuntimeError("o arquivo baixado não é uma planilha .xlsx válida")
    if saida.exists() and conteudo(saida.read_bytes()) == novo_conteudo:
        print(f"{serie['id']}: planilha sem mudanças.")
        return
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_bytes(novo)
    print(f"{serie['id']}: planilha atualizada ({len(novo)} bytes).")


falhas = 0
for serie in SERIES:
    if not serie.get("drive"):
        continue
    try:
        baixar(serie)
    except Exception as e:  # uma série com problema não impede as outras
        falhas += 1
        print(f"{serie['id']}: falhou ({e})")
if falhas == len([s for s in SERIES if s.get("drive")]):
    raise SystemExit("Nenhuma planilha pôde ser baixada.")
