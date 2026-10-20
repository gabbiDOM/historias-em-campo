"""Lê o feed público do canal no YouTube e grava dados/videos.json com os episódios da série."""
import json
import os
import pathlib
import re
import urllib.request
import xml.etree.ElementTree as ET

CHANNEL_ID = os.environ.get("CHANNEL_ID", "UCOvUI__POkLL10CgCplahQw")
# Só entram vídeos cujo título contém este texto.
FILTRO = os.environ.get("FILTRO", "Modo Carreira Realista Wolfsburg")
# Vídeo sem "EP NN" cujo título começa com este texto vira a INTRO (episódio 0).
INTRO = os.environ.get("INTRO", "Histórias em Campo")
SAIDA = pathlib.Path("dados/videos.json")
NS = {
    "a": "http://www.w3.org/2005/Atom",
    "yt": "http://www.youtube.com/xml/schemas/2015",
}

url = f"https://www.youtube.com/feeds/videos.xml?channel_id={CHANNEL_ID}"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
root = ET.fromstring(urllib.request.urlopen(req, timeout=30).read())

# O feed só traz os 15 vídeos mais recentes, então os já conhecidos são mantidos.
antigos = json.loads(SAIDA.read_text("utf-8")).get("videos", []) if SAIDA.exists() else []
videos = {v["id"]: v for v in antigos}

for entry in root.findall("a:entry", NS):
    titulo = entry.find("a:title", NS).text or ""
    if FILTRO.lower() not in titulo.lower():
        continue
    vid = entry.find("yt:videoId", NS).text
    m = re.search(r"\bEP\s*0*(\d+)", titulo, re.IGNORECASE)
    if m:
        episodio = int(m.group(1))
    elif titulo.lower().startswith(INTRO.lower()):
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
SAIDA.write_text(json.dumps({"canal": CHANNEL_ID, "videos": lista}, ensure_ascii=False, indent=2) + "\n", "utf-8")
print(f"{len(lista)} vídeo(s) da série em {SAIDA}")
