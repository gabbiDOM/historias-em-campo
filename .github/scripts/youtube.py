"""Busca os vídeos do canal e grava dados/series/<id>/videos.json com os episódios
de cada série cadastrada em dados/series.json.

Fonte principal: o feed público do canal. Quando o feed do YouTube está fora do ar
(acontece de responder 404 por horas, para todos os canais), lê a página de vídeos do
canal e a página de cada vídeo recente."""
import html
import json
import os
import pathlib
import re
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

CHANNEL_ID = os.environ.get("CHANNEL_ID", "UCOvUI__POkLL10CgCplahQw")
SERIES = json.loads(pathlib.Path("dados/series.json").read_text("utf-8"))["series"]
NS = {
    "a": "http://www.w3.org/2005/Atom",
    "yt": "http://www.youtube.com/xml/schemas/2015",
}
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9",
}
RECENTES = 15  # quantos vídeos mais novos a página de reserva consulta


def baixar(url):
    return urllib.request.urlopen(urllib.request.Request(url, headers=HEADERS), timeout=30).read()


def via_feed():
    """Lista de (id, titulo, publicado) pelo feed; tenta algumas vezes antes de desistir."""
    url = f"https://www.youtube.com/feeds/videos.xml?channel_id={CHANNEL_ID}"
    for tentativa in range(1, 4):
        try:
            raiz = ET.fromstring(baixar(url))
            return [(e.find("yt:videoId", NS).text, e.find("a:title", NS).text or "", e.find("a:published", NS).text)
                    for e in raiz.findall("a:entry", NS)]
        except Exception as e:
            print(f"feed do YouTube indisponível (tentativa {tentativa}/3): {e}")
            time.sleep(10 * tentativa)
    return None


def via_pagina():
    """Lista de (id, titulo, publicado) lendo a aba Vídeos do canal e a página de cada vídeo."""
    pagina = baixar(f"https://www.youtube.com/channel/{CHANNEL_ID}/videos").decode("utf-8", "ignore")
    ids = list(dict.fromkeys(re.findall(r'"videoId":"([\w-]{11})"', pagina)))[:RECENTES]
    videos = []
    for vid in ids:
        try:
            p = baixar(f"https://www.youtube.com/watch?v={vid}").decode("utf-8", "ignore")
            titulo = re.search(r'<meta name="title" content="([^"]*)"', p)
            data = re.search(r'"publishDate":"([^"]+)"', p) or re.search(r'itemprop="datePublished" content="([^"]+)"', p)
            if titulo and data:
                videos.append((vid, html.unescape(titulo.group(1)), data.group(1)))
        except Exception as e:
            print(f"não consegui ler o vídeo {vid}: {e}")
    return videos


entradas = via_feed()
if entradas is None:
    try:
        entradas = via_pagina()
        print(f"usando a página do canal como reserva: {len(entradas)} vídeo(s) lidos")
    except Exception as e:
        print(f"página do canal também indisponível: {e}")
if not entradas:
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
    # Só os vídeos mais recentes vêm de cada fonte, então os já conhecidos são mantidos.
    antigos = json.loads(saida.read_text("utf-8")).get("videos", []) if saida.exists() else []
    videos = {v["id"]: v for v in antigos}

    for vid, titulo, publicado in entradas:
        if filtro.lower() not in titulo.lower():
            continue
        m = re.search(r"\bEP\s*0*(\d+)", titulo, re.IGNORECASE)
        if m:
            episodio = int(m.group(1))
        elif intro and titulo.lower().startswith(intro.lower()):
            episodio = 0
        else:
            episodio = None
        # Data sempre em UTC ("2026-10-27T21:00:21+00:00"), igual ao feed.
        try:
            publicado = datetime.fromisoformat(publicado).astimezone(timezone.utc).isoformat()
        except ValueError:
            pass
        videos[vid] = {"id": vid, "titulo": titulo, "publicado": publicado, "episodio": episodio}

    lista = sorted(videos.values(), key=lambda v: v["publicado"])
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(json.dumps({"canal": CHANNEL_ID, "videos": lista}, ensure_ascii=False, indent=2) + "\n", "utf-8")
    print(f"{serie['id']}: {len(lista)} vídeo(s) da série")
