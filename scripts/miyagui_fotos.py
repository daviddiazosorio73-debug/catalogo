"""Fotos de industriasmiyagui.com (proveedor) para los productos del catálogo que no tienen foto.

Lee las fichas /equipo/ del sitio, reconoce el modelo, lo cruza con los códigos del catálogo y guarda
la foto principal en assets/m/<id>.webp. Resultado en assets/m/miyagui.json (no borra nada).
"""
import io, json, re, time
from pathlib import Path
import requests
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "m"; OUT.mkdir(parents=True, exist_ok=True)
SITE = "https://industriasmiyagui.com"
H = {"User-Agent": "Mozilla/5.0 (TecnoLogIA catalogo; contacto WhatsApp 928146619)"}
api = re.search(r"api:\s*'([^']+)'", (ROOT / "config.js").read_text()).group(1)
N = lambda s: re.sub(r"[^A-Z0-9]", "", str(s).upper())


def get(url, **kw):
    for i in range(3):
        try:
            r = requests.get(url, headers=H, timeout=40, **kw)
            if r.status_code == 200:
                return r
        except Exception:
            pass
        time.sleep(2 + 3 * i)
    return None


def urls_equipos():
    urls = set()
    for sm in ["/wp-sitemap.xml", "/sitemap_index.xml", "/sitemap.xml"]:
        r = get(SITE + sm)
        if not r:
            continue
        for loc in re.findall(r"<loc>([^<]+)</loc>", r.text):
            if loc.endswith(".xml") and re.search(r"equipo|product|posts", loc):
                r2 = get(loc)
                if r2:
                    urls |= {u for u in re.findall(r"<loc>([^<]+)</loc>", r2.text) if "/equipo/" in u}
            elif "/equipo/" in loc:
                urls.add(loc)
        if urls:
            break
    if not urls:  # respaldo: recorrer catálogo y categorías
        cola, vistos = [SITE + "/catalogo/"], set()
        while cola and len(vistos) < 120:
            u = cola.pop(0)
            if u in vistos:
                continue
            vistos.add(u)
            r = get(u)
            if not r:
                continue
            for h in re.findall(r'href="(https://industriasmiyagui\.com/[^"#?]+)"', r.text):
                if "/equipo/" in h:
                    urls.add(h)
                elif re.search(r"/(linea-[a-z]+|catalogo|marca)/", h) and h not in vistos:
                    cola.append(h)
            for pag in re.findall(r'href="(https://industriasmiyagui\.com/catalogo/page/\d+/)"', r.text):
                cola.append(pag)
            time.sleep(0.8)
    return sorted(urls)


def ficha(url):
    r = get(url)
    if not r:
        return None
    t = r.text
    titulo = re.sub(r"\s+", " ", (re.search(r"<title>(.*?)</title>", t, re.S) or [None, ""])[1]).strip()
    og = re.search(r'property="og:image"\s+content="([^"]+)"', t) or re.search(r'content="([^"]+)"\s+property="og:image"', t)
    imgs = [og.group(1)] if og else []
    imgs += re.findall(r'(https://industriasmiyagui\.com/wp-content/uploads/[^"\'\s>]+\.(?:jpe?g|png|webp))', t, re.I)
    imgs = [i for i in dict.fromkeys(imgs) if not re.search(r"logo|icon|favicon|-\d{2,3}x\d{2,3}\.", i, re.I)]
    texto = re.sub(r"<[^>]+>", " ", t)
    modelo = re.findall(r"(?:Modelo|Código|SKU)\s*:?\s*([A-Z0-9][A-Z0-9\-\./ ]{2,25})", texto)
    return {"url": url, "titulo": titulo, "modelos": modelo[:3], "imgs": imgs[:6], "clave": (url.rstrip("/").rsplit("/", 1)[-1] + " " + titulo + " " + " ".join(modelo[:3])).upper()}


def main():
    prods = requests.get(api, params={"accion": "productos"}, headers=H, timeout=120).json().get("productos", [])
    sin = [p for p in prods if not p.get("foto")]
    print("productos sin foto:", len(sin))
    eq = urls_equipos(); print("fichas en el sitio:", len(eq))
    fichas = []
    for u in eq:
        f = ficha(u)
        if f:
            fichas.append(f)
        time.sleep(0.8)
    res, prev = {}, {}
    pf = OUT / "miyagui.json"
    if pf.exists():
        prev = json.loads(pf.read_text())
    for p in sin:
        c = N(p["codigo"])
        if len(c) < 4:
            continue
        # código exacto (con o sin guiones/espacios) y que no siga otra letra o número: BC-2900 ≠ BC-2900FC, DK-800 ≠ DK-8001
        pat = r"(?<![A-Z0-9])" + r"[\s\-_./]?".join(map(re.escape, c)) + r"(?![A-Z0-9])"
        cand = [f for f in fichas if re.search(pat, f["clave"])]
        if not cand or not cand[0]["imgs"]:
            continue
        f = cand[0]
        dest = OUT / (re.sub(r"[^A-Za-z0-9_-]", "_", p["id"]) + ".webp")
        if not dest.exists():
            ok = False
            for im in f["imgs"]:
                r = get(im)
                if not r:
                    continue
                try:
                    img = Image.open(io.BytesIO(r.content))
                    if max(img.size) < 300:
                        continue
                    img = img.convert("RGBA")
                    fondo = Image.new("RGB", img.size, (255, 255, 255)); fondo.paste(img, mask=img.split()[3])
                    fondo.thumbnail((1000, 1000)); fondo.save(dest, "WEBP", quality=82, method=5); ok = True; break
                except Exception:
                    continue
            if not ok:
                continue
        res[p["id"]] = {"codigo": p["codigo"], "archivo": "assets/m/" + dest.name, "pagina": f["url"], "titulo": f["titulo"]}
    prev.update(res)
    pf.write_text(json.dumps(prev, ensure_ascii=False, indent=1))
    (OUT / "fichas_sitio.json").write_text(json.dumps([{k: f[k] for k in ("url", "titulo", "modelos")} for f in fichas], ensure_ascii=False, indent=1))
    print("fotos asignadas:", len(res), "· total acumulado:", len(prev))


if __name__ == "__main__":
    main()
