# Uso: CA_BUNDLE=/root/.ccr/ca-bundle.crt python3 herramientas/actualizar_fitos.py fitos.json /tmp/fichas
# Necesita: pip install pdfplumber. Con REUSAR=fitos.json reutiliza la lista de productos por cultivo.
"""Construye fitos.json: productos fitosanitarios autorizados en los cultivos de la explotación,
con sus usos (cultivo, agente, dosis) leídos de la ficha oficial del Registro del MAPA."""
import json, re, ssl, sys, os, time, unicodedata, urllib.request, urllib.parse, io, concurrent.futures as cf
import pdfplumber

B = 'https://servicio.mapa.gob.es/regfiweb'
CTX = ssl.create_default_context(cafile=os.environ.get('CA_BUNDLE') or None)
OUT = sys.argv[1]
CACHE = sys.argv[2]
os.makedirs(CACHE, exist_ok=True)

# Cultivos del registro que interesan (nombre exacto del registro)
CULTIVOS = ['Trigo', 'Cebada', 'Avena', 'Centeno', 'Triticale', 'Cereales de invierno / primavera',
            'Lenteja', 'Garbanzo', 'Guisante para grano', 'Almorta', 'Leguminosas de grano',
            'Yeros', 'Veza', 'Esparceta', 'Alfalfa', 'Leguminosas forrajeras',
            'Anís, matalauva', 'Cilantro', 'Girasol', 'Colza', 'Camelina', 'Barbechos', 'Chopos y álamos']


def get(url, data=None, binary=False, tries=4):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, data=urllib.parse.urlencode(data).encode() if data else None)
            with urllib.request.urlopen(req, timeout=120, context=CTX) as r:
                b = r.read()
            return b if binary else b.decode('utf-8', 'replace')
        except Exception as e:
            err = e; time.sleep(2 + 3 * i)
    raise err


def norm(s):
    return re.sub(r'\s+', ' ', unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode().lower()).strip()


def id_cultivo(nombre):
    h = get(f'{B}/Cultivos/CultivosGrid?' + urllib.parse.urlencode({'nombreComun': nombre.split(',')[0].split('/')[0].strip()}))
    import html as H
    for m in re.finditer(r'data-nombre="([^"]*)"\s*data-id="(\d+)"', h):
        if H.unescape(m.group(1)).strip() == nombre: return m.group(2)
    raise SystemExit('No encuentro el cultivo ' + nombre)


def productos_de(idc):
    ids = set(); page = 1
    while True:
        h = get(f'{B}/Cultivos/ProductosGrid?IdCultivo={idc}&productosGrid-page={page}')
        nuevos = set(re.findall(r'data-id="(\d+)"', h))
        tot = re.search(r'de un total de (\d+)', h)
        ids |= nuevos
        if not nuevos or not tot or page * 5 >= int(tot.group(1)): break
        page += 1
    return ids


DOSIS = re.compile(r'(\d+(?:[.,]\d+)?)(?:\s*-\s*(\d+(?:[.,]\d+)?))?\s*(cc|ml|l|g|kg|gr)\s*/\s*(ha|hl|100 ?l|kg|t|m2|100 ?m2|planta|árbol|arbol)', re.I)


def leer_usos(pdf_bytes):
    usos = []  # [cultivo, agente, texto dosis]
    ultimo = None
    ultimo_cols = None
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        en_usos = False
        for pg in pdf.pages:
            ws = pg.extract_words(keep_blank_chars=False, use_text_flow=False)
            txt = ' '.join(w['text'] for w in ws)
            # localizar la cabecera USO / AGENTE / Dosis
            cab = None
            for w in ws:
                if w['text'] == 'AGENTE':
                    fila = [x for x in ws if abs(x['top'] - w['top']) < 3]
                    nombres = [x['text'] for x in fila]
                    if 'USO' in nombres and 'Dosis' in nombres:
                        cab = {x['text']: x for x in fila}; break
            if not cab and not en_usos: continue
            if cab:
                en_usos = True
                y0 = cab['AGENTE']['bottom']
                ht, hb = cab['AGENTE']['top'], cab['AGENTE']['bottom']
                xs = sorted({round(l['x0']) for l in pg.lines if abs(l['x0'] - l['x1']) < 1 and l['top'] <= ht + 1 and l['bottom'] >= hb - 1})
                xs2 = []
                for x in xs:
                    if not xs2 or x - xs2[-1] > 4: xs2.append(x)
                fila = [x for x in ws if abs(x['top'] - ht) < 3]
                def col_de(palabra, defecto):
                    w = next((x for x in fila if x['text'].startswith(palabra)), None)
                    if not w or len(xs2) < 3: return defecto
                    for a0, b0 in zip(xs2, xs2[1:]):
                        if a0 - 2 <= w['x0'] < b0: return (a0, b0)
                    return defecto
                ini = xs2[0] - 2 if xs2 else cab['USO']['x0'] - 40
                ca = col_de('AGENTE', (cab['AGENTE']['x0'] - 25, cab['Dosis']['x0'] - 40))
                cd = col_de('Dosis', (cab['Dosis']['x0'] - 40, cab['Dosis']['x0'] + 80))
                cc = col_de('Condic', (cd[1], pg.width))
                cols = (ini, ca[0], cd[0], cd[1], cc[0], cc[1])
            elif ultimo_cols:
                cols = ultimo_cols; y0 = 0
            ultimo_cols = cols
            xu0, xa, xd, xc, xk0, xk1 = cols
            # fin de la tabla: «Plazos de Seguridad» o «Condiciones Generales»
            fin = min([w['top'] for w in ws if w['top'] > y0 and w['text'] in ('Plazos', 'Condiciones', 'Clase', 'Mitigación')] + [pg.height])
            # separadores de bloque en la columna USO
            seps = sorted({round(l['top']) for l in pg.lines if l['top'] > y0 + 2 and l['top'] < fin + 2 and l['x0'] <= xa - 20 and l['x1'] >= xa - 5 and abs(l['top'] - l['bottom']) < 1})
            limites = [y0] + [s for s in seps if s > y0 + 2] + [fin]
            limites = sorted(set(limites))
            for a, b in zip(limites, limites[1:]):
                if b - a < 4: continue
                en = [w for w in ws if a - 1 <= w['top'] < b - 1]
                uso = ' '.join(w['text'] for w in en if w['x0'] < xa - 1 and w['x0'] >= xu0)
                ag_lin, do_lin, co_lin = {}, {}, {}
                for w in en:
                    k = round(w['top'] / 4)
                    if xa - 2 <= w['x0'] < xd: ag_lin.setdefault(k, []).append(w['text'])
                    elif xd - 2 <= w['x0'] < xc: do_lin.setdefault(k, []).append(w['text'])
                    elif xk0 - 2 <= w['x0'] < xk1: co_lin.setdefault(k, []).append(w['text'])
                uso = uso.strip() or (ultimo or '')
                if not uso: continue
                ultimo = uso
                # emparejar cada dosis con el agente más cercano por arriba/abajo
                ags = sorted((k, ' '.join(v)) for k, v in ag_lin.items())
                if not any(DOSIS.search(' '.join(v)) for v in do_lin.values()):
                    # fichas nuevas: «Dosis máxima por aplicación 1,5 l/ha» en las condiciones
                    texto = ' '.join(' '.join(v) for k, v in sorted(co_lin.items()))
                    for m in re.finditer(r'[Dd]osis\s+m[aá]xima[^.;]{0,60}?(\d+(?:[.,]\d+)?\s*(?:cc|ml|l|g|kg)\s*/\s*ha)', texto):
                        usos.append([uso, (ags[0][1] if ags else ''), 'MAX ' + m.group(1)])
                for k, v in sorted(do_lin.items()):
                    d = ' '.join(v)
                    if not DOSIS.search(d): continue
                    cerca = min(ags, key=lambda x: abs(x[0] - k))[1] if ags else ''
                    usos.append([uso, cerca, d])
            if 'Plazos de Seguridad' in txt or 'Condiciones Generales' in txt: en_usos = False if not cab or fin < pg.height else en_usos
    return usos


def num(s): return float(s.replace(',', '.'))


def main():
    prod = json.loads(get(f'{B}/Exportaciones/ExportJsonProductos', {'dataDto': json.dumps({'nombreComercial': '', 'titular': '', 'numRegistro': '', 'fabricante': '', 'idSustancia': None, 'idPlaga': None, 'idFuncion': None, 'idEstado': None, 'idAmbito': None, 'idCultivo': None})}))
    if isinstance(prod, str): prod = json.loads(prod)
    fecha = prod.get('Fecha'); prod = {str(p['IdProducto']): p for p in json.loads(prod['Contenido'])}
    print('productos en registro', len(prod), fecha, flush=True)
    por_cultivo = {}
    previo = os.environ.get('REUSAR')  # reutiliza la lista de productos por cultivo de un fitos.json anterior
    if previo:
        for p in json.load(open(previo))['productos']:
            for c in p['c']: por_cultivo.setdefault(c, set()).add(str(p['id']))
    else:
        for c in CULTIVOS:
            idc = id_cultivo(c); ids = productos_de(idc); por_cultivo[c] = ids
            print(c, idc, len(ids), flush=True)
    todos = sorted(set().union(*por_cultivo.values()), key=int)
    if len(sys.argv) > 3: todos = sys.argv[3].split(',')
    print('fichas a leer', len(todos), flush=True)
    relevantes = {norm(c) for c in CULTIVOS}

    def ficha(pid):
        ruta = os.path.join(CACHE, pid + '.pdf')
        if not os.path.exists(ruta) or os.path.getsize(ruta) < 1000:
            b = get(f'{B}/Productos/ExportFichaProductoPdfGet?idProducto={pid}', binary=True)
            with open(ruta, 'wb') as f: f.write(b)
        with open(ruta, 'rb') as f: b = f.read()
        try: return pid, leer_usos(b), None
        except Exception as e: return pid, [], str(e)

    salida = []; errores = 0
    with cf.ThreadPoolExecutor(6) as ex:
        for n, (pid, usos, err) in enumerate(ex.map(ficha, todos), 1):
            if err: errores += 1
            p = prod.get(pid, {})
            u = {}
            for uso, ag, d in usos:
                # una celda USO puede listar varios cultivos separados por comas
                for cu in [x.strip() for x in re.split(r',\s*(?=[A-ZÁÉÍÓÚ])', uso)]:
                    if norm(cu) not in relevantes: continue
                    for m in DOSIS.finditer(d):
                        a = num(m.group(1)); bmax = num(m.group(2)) if m.group(2) else a
                        if d.startswith('MAX '): a = 0.0  # solo se conoce la dosis máxima
                        un = (m.group(3).lower().replace('gr', 'g') + '/' + m.group(4).lower().replace(' ', ''))
                        u.setdefault(cu, []).append([ag[:60], a, bmax, un])
            salida.append({'id': int(pid), 'r': p.get('NumRegistro', ''), 'n': p.get('Nombre', ''), 'f': p.get('Formulado', ''), 'e': p.get('Estado', ''),
                           'cad': p.get('StrFechaCaducidad', ''), 'lv': p.get('StrFechaLimiteVenta', ''),
                           'c': sorted(c for c, ids in por_cultivo.items() if pid in ids), 'u': u})
            if n % 50 == 0: print('leídas', n, 'errores', errores, flush=True)
    json.dump({'fecha': fecha, 'fuente': 'Registro de Productos Fitosanitarios (MAPA)', 'cultivos': CULTIVOS, 'productos': salida}, open(OUT, 'w'), ensure_ascii=False, separators=(',', ':'))
    print('OK', len(salida), 'productos; errores', errores, 'tamaño', os.path.getsize(OUT), flush=True)


if __name__ == '__main__':
    main()
