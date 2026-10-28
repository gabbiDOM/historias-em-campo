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


UNIDADES = {  # "há 5 h", "há 2 dias", "3 weeks ago"... em segundos
    "s": 1, "seg": 1, "second": 1, "min": 60, "minute": 60, "h": 3600, "hora": 3600, "hour": 3600,
    "dia": 86400, "day": 86400, "semana": 604800, "week": 604800, "mes": 2592000, "mês": 2592000,
    "meses": 2592000, "month": 2592000, "ano": 31536000, "year": 31536000,
}


def data_relativa(texto):
    """Converte "há 5 h" / "5 hours ago" numa data UTC aproximada (o feed traz a exata)."""
    m = re.search(r"(\d+)\s*([a-zA-Zêç]+)", texto)
    if not m:
        return None
    unidade = m.group(2).lower().rstrip("s") if m.group(2).lower() not in ("s", "meses") else m.group(2).lower()
    segundos = UNIDADES.get(unidade) or UNIDADES.get(m.group(2).lower())
    if not segundos:
        return None
    return datetime.fromtimestamp(time.time() - int(m.group(1)) * segundos, timezone.utc).isoformat(timespec="seconds")


def via_pagina():
    """Lista de (id, titulo, publicado) lendo a aba Vídeos do canal (título completo e data relativa)."""
    pagina = baixar(f"https://www.youtube.com/channel/{CHANNEL_ID}/videos").decode("utf-8", "ignore")
    videos = []
    for bloco in pagina.split('"lockupViewModel":{"contentImage"')[1:RECENTES + 1]:
        vid = re.search(r"/vi/([\w-]{11})/", bloco)
        titulo = re.search(r'"lockupMetadataViewModel":\{"title":\{"content":"((?:[^"\\]|\\.)*)"', bloco)
        quando = re.search(r'"content":"((?:há|ha) [^"]+|[^"]* ago)"', bloco)
        if not (vid and titulo):
            continue
        texto = html.unescape(json.loads(f'"{titulo.group(1)}"'))
        videos.append((vid.group(1), texto, data_relativa(quando.group(1)) if quando else None))
    return videos


entradas = via_feed()
data_exata = entradas is not None
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
        # A página do canal só dá data aproximada ("há 5 h"): para vídeo já conhecido, vale a data salva.
        anterior = videos.get(vid, {}).get("publicado")
        if not data_exata and anterior:
            publicado = anterior
        publicado = publicado or anterior or datetime.now(timezone.utc).isoformat(timespec="seconds")
        try:  # data sempre em UTC ("2026-10-27T21:00:21+00:00"), igual ao feed
            publicado = datetime.fromisoformat(publicado).astimezone(timezone.utc).isoformat()
        except ValueError:
            pass
        videos[vid] = {"id": vid, "titulo": titulo, "publicado": publicado, "episodio": episodio}

    lista = sorted(videos.values(), key=lambda v: v["publicado"])
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(json.dumps({"canal": CHANNEL_ID, "videos": lista}, ensure_ascii=False, indent=2) + "\n", "utf-8")
    print(f"{serie['id']}: {len(lista)} vídeo(s) da série")
