"""Lê o feed público do canal no YouTube e grava dados/series/<id>/videos.json
com os episódios de cada série cadastrada em dados/series.json."""
import json
import os
import pathlib
import re
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET

CHANNEL_ID = os.environ.get("CHANNEL_ID", "UCOvUI__POkLL10CgCplahQw")
SERIES = json.loads(pathlib.Path("dados/series.json").read_text("utf-8"))["series"]
NS = {
    "a": "http://www.w3.org/2005/Atom",
    "yt": "http://www.youtube.com/xml/schemas/2015",
}

url = f"https://www.youtube.com/feeds/videos.xml?channel_id={CHANNEL_ID}"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
# O feed do YouTube às vezes responde 404/5xx por alguns minutos. Tenta algumas vezes e,
# se continuar fora, mantém os vídeos já conhecidos em vez de derrubar a atualização inteira.
entradas = None
for tentativa in range(1, 5):
    try:
        entradas = ET.fromstring(urllib.request.urlopen(req, timeout=30).read()).findall("a:entry", NS)
        break
    except Exception as e:
        print(f"feed do YouTube indisponível (tentativa {tentativa}/4): {e}")
        time.sleep(15 * tentativa)
if entradas is None:
    print("Seguindo sem novidades do YouTube; os vídeos já conhecidos continuam no site.")
    sys.exit(0)

for serie in SERIES:
    # Só entram vídeos cujo título contém o filtro da série.
    filtro = serie.get("youtube_filtro") or ""
    # Vídeo sem "EP NN" cujo título começa com este texto vira a INTRO (episódio 0).
    intro = serie.get("youtube_intro") or ""
    if not filtro:
        continue
    saida = pathlib.Path(f"dados/series/{serie['id']}/videos.json")
    # O feed só traz os 15 vídeos mais recentes, então os já conhecidos são mantidos.
    antigos = json.loads(saida.read_text("utf-8")).get("videos", []) if saida.exists() else []
    videos = {v["id"]: v for v in antigos}

    for entry in entradas:
        titulo = entry.find("a:title", NS).text or ""
        if filtro.lower() not in titulo.lower():
            continue
        vid = entry.find("yt:videoId", NS).text
        m = re.search(r"\bEP\s*0*(\d+)", titulo, re.IGNORECASE)
        if m:
            episodio = int(m.group(1))
        elif intro and titulo.lower().startswith(intro.lower()):
            episodio = 0
        else:
            episodio = None
        videos[vid] = {
            "id": vid,
            "titulo": titulo,
            "publicado": entry.find("a:published", NS).text,
            "episodio": episodio,
        }

    lista = sorted(videos.values(), key=lambda v: v["publicado"])
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(json.dumps({"canal": CHANNEL_ID, "videos": lista}, ensure_ascii=False, indent=2) + "\n", "utf-8")
    print(f"{serie['id']}: {len(lista)} vídeo(s) da série")
