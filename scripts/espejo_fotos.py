"""Copia las fotos de Google Drive del catálogo al sitio (assets/d/<id>.webp).

Por qué: Google Drive limita la cantidad de imágenes que entrega seguidas a un mismo visitante,
y algunas fotos no cargan en el celular. Con la copia en el sitio cargan rápido y siempre.
No borra nada: solo agrega las fotos nuevas. Deja en assets/d/fallas.json las que no se pudieron leer.
"""
import io, json, re, sys, time
from pathlib import Path
import requests
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
D = ROOT / "assets" / "d"
D.mkdir(parents=True, exist_ok=True)
api = re.search(r"api:\s*'([^']+)'", (ROOT / "config.js").read_text()).group(1)
H = {"User-Agent": "Mozilla/5.0 (catalogo-espejo)"}


def ids_del_catalogo():
    d = requests.get(api, params={"accion": "productos"}, headers=H, timeout=120).json()
    ids, prod = [], {}
    for p in d.get("productos", []):
        for u in [p.get("foto")] + list(p.get("fotos") or []):
            m = re.search(r"/d/([\w-]{20,})|[?&]id=([\w-]{20,})", u or "")
            if m:
                i = m.group(1) or m.group(2)
                if i not in prod:
                    ids.append(i); prod[i] = p.get("codigo", "")
    return ids, prod


def bajar(i):
    urls = [f"https://lh3.googleusercontent.com/d/{i}=w1000",
            f"https://drive.google.com/thumbnail?id={i}&sz=w1000",
            f"https://drive.google.com/uc?export=download&id={i}"]
    ultimo = ""
    for u in urls:
        for intento in range(3):
            try:
                r = requests.get(u, headers=H, timeout=60)
            except Exception as e:
                ultimo = "red: " + str(e)[:80]; time.sleep(2); continue
            if r.status_code == 429:
                time.sleep(10 * (intento + 1)); ultimo = "429 límite de Google"; continue
            tipo = r.headers.get("content-type", "")
            if r.status_code == 200 and tipo.startswith("image/"):
                return r.content, ""
            ultimo = f"{r.status_code} {tipo.split(';')[0] or 'sin tipo'}"
            break
    return None, ultimo


def main():
    ids, prod = ids_del_catalogo()
    fallas, nuevas = {}, 0
    for n, i in enumerate(ids):
        dest = D / f"{i}.webp"
        if dest.exists():
            continue
        data, err = bajar(i)
        if data is None:
            fallas[i] = {"codigo": prod[i], "error": err}
            continue
        try:
            im = Image.open(io.BytesIO(data))
            im = im.convert("RGBA") if im.mode in ("P", "LA", "RGBA") else im.convert("RGB")
            if im.mode == "RGBA":  # fondo blanco
                fondo = Image.new("RGB", im.size, (255, 255, 255)); fondo.paste(im, mask=im.split()[3]); im = fondo
            im.thumbnail((1000, 1000))
            im.save(dest, "WEBP", quality=80, method=5)
            nuevas += 1
        except Exception as e:
            fallas[i] = {"codigo": prod[i], "error": "no es imagen: " + str(e)[:60]}
        time.sleep(0.4)
        if n % 50 == 0:
            print(n, "/", len(ids), "nuevas", nuevas, "fallas", len(fallas), flush=True)
    presentes = sorted(p.stem for p in D.glob("*.webp"))
    (D / "index.json").write_text(json.dumps(presentes))
    (D / "fallas.json").write_text(json.dumps(fallas, ensure_ascii=False, indent=1))
    print(f"Listo: {len(presentes)} fotos en el sitio · nuevas {nuevas} · fallas {len(fallas)}")


if __name__ == "__main__":
    main()
