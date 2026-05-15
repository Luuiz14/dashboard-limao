"""Dashboard Limao - Servidor Cloud (Render.com)"""

import io, json, os, re, sys, base64
from datetime import datetime
from collections import Counter, defaultdict

# ── CONFIGURACAO ──────────────────────────────────────────────────────────────
# Link de compartilhamento do OneDrive (qualquer pessoa com o link pode visualizar)
ONEDRIVE_SHARE_URL = "https://1drv.ms/x/c/e35c45354ce94f38/IQC1CiXP4Ai3TbEJ3GMvbdI7AX_ekay8pgXBZyB4InVfMLE?e=Erlyde"

META_VENDAS = 29196
PORT        = int(os.environ.get("PORT", 8050))
# ──────────────────────────────────────────────────────────────────────────────

try:
    from flask import Flask, jsonify, send_from_directory
    import openpyxl
    import requests
except ImportError:
    print("Execute: pip install flask openpyxl requests gunicorn")
    sys.exit(1)

app = Flask(__name__, static_folder=".", static_url_path="")


ONEDRIVE_RESID   = "E35C45354CE94F38!scf250ab508e04db7b109dc632f6dd23b"
ONEDRIVE_CID     = "e35c45354ce94f38"
ONEDRIVE_AUTHKEY = "!Erlyde"
ONEDRIVE_PATH    = "Vendas%20e%20Renova%c3%a7%c3%b5es%20-%20Marginal.xlsx"


def baixar_planilha():
    """Baixa a planilha do OneDrive usando multiplas estrategias."""
    UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    sess = requests.Session()
    sess.headers.update({"User-Agent": UA})

    b64 = base64.urlsafe_b64encode(ONEDRIVE_SHARE_URL.encode()).rstrip(b"=").decode()
    erros = []
    html_trecho = ""

    # Estrategia 1: Download via resid + authkey (formato SPO migrado)
    try:
        resid_enc = ONEDRIVE_RESID.replace("!", "%21")
        auth_enc  = ONEDRIVE_AUTHKEY.replace("!", "%21")
        url = f"https://onedrive.live.com/download?resid={resid_enc}&authkey={auth_enc}"
        r = sess.get(url, timeout=30, allow_redirects=True)
        if r.ok and r.content[:4] == b"PK\x03\x04":
            return io.BytesIO(r.content)
        erros.append(f"ResidDL={r.status_code} url={r.url[:60]}")
    except Exception as e:
        erros.append(f"ResidDL_exc={e}")

    # Estrategia 2: Download via caminho do arquivo + authkey
    try:
        url = (f"https://onedrive.live.com/personal/{ONEDRIVE_CID}/"
               f"Documents/Documentos/{ONEDRIVE_PATH}"
               f"?ga=1&authkey={ONEDRIVE_AUTHKEY.replace('!', '%21')}")
        r = sess.get(url, timeout=30, allow_redirects=True)
        if r.ok and r.content[:4] == b"PK\x03\x04":
            return io.BytesIO(r.content)
        erros.append(f"PathDL={r.status_code}")
    except Exception as e:
        erros.append(f"PathDL_exc={e}")

    # Estrategia 3: Microsoft Graph API
    try:
        url = f"https://graph.microsoft.com/v1.0/shares/u!{b64}/driveItem/content"
        r = sess.get(url, timeout=30, allow_redirects=True)
        if r.ok and r.content[:4] == b"PK\x03\x04":
            return io.BytesIO(r.content)
        erros.append(f"Graph={r.status_code}")
    except Exception as e:
        erros.append(f"Graph_exc={e}")

    # Estrategia 4: Seguir share link e extrair URL do HTML da pagina
    try:
        r = sess.get(ONEDRIVE_SHARE_URL, timeout=30, allow_redirects=True)
        if r.ok and r.content[:4] == b"PK\x03\x04":
            return io.BytesIO(r.content)
        html_trecho = r.text[:500].replace("\n", " ")
        for pat in [
            r'"[Ff]ile[Gg]et[Uu]rl"\s*:\s*"(https://[^"]+)"',
            r'"[Ff]ile[Gg]et[Uu]rl"\s*:\s*"(https:\\/\\/[^"]+)"',
            r'"[Dd]ownload[Uu]rl"\s*:\s*"(https://[^"]+)"',
            r'"[Dd]ownload[Uu]rl"\s*:\s*"(https:\\/\\/[^"]+)"',
            r'data-download-url="([^"]+)"',
            r'"url"\s*:\s*"(https://[^"]+\.xlsx[^"]*)"',
        ]:
            m = re.search(pat, r.text)
            if m:
                dl = m.group(1).replace("\\/", "/").replace("\\u0026", "&")
                r2 = sess.get(dl, timeout=30, allow_redirects=True)
                if r2.ok and r2.content[:4] == b"PK\x03\x04":
                    return io.BytesIO(r2.content)
        erros.append(f"HTML={r.status_code} url={r.url[:60]}")
    except Exception as e:
        erros.append(f"HTML_exc={e}")

    raise Exception("Falha: " + " | ".join(erros) + " || HTML=" + html_trecho)


def ler_planilha():
    """Baixa e parseia os dados das abas de Maio."""
    try:
        arquivo = baixar_planilha()
        wb = openpyxl.load_workbook(arquivo, data_only=True)
    except Exception as e:
        return {"erro": "Nao foi possivel acessar a planilha: " + str(e)}

    resultado = {}

    # VENDAS MAIO
    try:
        ws = wb["Vendas Maio"]
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        col = {str(h).strip(): i for i, h in enumerate(header) if h}

        vendas = []
        daily = defaultdict(float)

        for row in rows[1:]:
            nome = row[col.get("Nome", 1)] if col.get("Nome") is not None else None
            if not nome:
                continue
            tipo = row[col.get("Venda", 2)]
            data = row[col.get("Dia que fechou", 3)]
            val  = row[col.get("Valor medio", col.get("Valor médio", 11))]
            resp = row[col.get("Responsavel", col.get("Responsável", 13))]
            if isinstance(val, str):
                val = None
            val = float(val) if val else 0.0
            vendas.append({
                "nome":  str(nome),
                "tipo":  str(tipo).strip() if tipo else "",
                "valor": val,
                "resp":  str(resp).strip() if resp else "",
                "dia":   data.day if isinstance(data, datetime) else None,
            })
            if isinstance(data, datetime):
                daily[data.day] += val

        total_fat  = sum(v["valor"] for o in vendas)
        total_qtd  = len(vendas)
        ticket_med = round(total_fat / total_qtd) if total_qtd else 0
        tipos    = Counter(v["tipo"] for v in vendas if v["tipo"])
        resp_qtd = Counter(v["resp"] for v in vendas if v["resp"])
        resp_fat = defaultdict(float)
        for v in vendas:
            if v["resp"]:
                resp_fat[v["resp"]] += v["valor"]

        hoje = datetime.today().day
        acum = []
        cum = 0.0
        for d in range(1, 32):
            cum += daily.get(d, 0)
            acum.append(round(cum) if (cum > 0 and d <= hoje) else None)

        vendedores = sorted(
            [{"nome": k, "qtd": v, "fat": round(resp_fat[k])} for k, v in resp_qtd.items()],
            key=lambda x: -x["fat"]
        )

        resultado["vendas"] = {
            "total_qtd":  total_qtd,
            "total_fat":  round(total_fat),
            "ticket_med": ticket_med,
            "meta":       META_VENDAS,
            "pct_meta":   round(total_fat / META_VENDAS * 100, 1),
            "tipos":      dict(tipos),
            "acumulado":  acum,
            "vendedores": vendedores,
        }
    except Exception as e:
        resultado["vendas"] = {"erro": str(e)}

    # RENOVACOES MAIO
    try:
        # Busca aba por palavra-chave (robusto a encoding de Unicode)
        _sn = next((n for n in wb.sheetnames if "Renova" in n and "Maio" in n), None)
        if not _sn:
            raise KeyError(f"Aba Renovacoes nao encontrada. Sheets: {wb.sheetnames}")
        ws = wb[_sn]
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        col = {str(h).strip(): i for i, h in enumerate(header) if h}

        sit_count  = Counter()
        sit_fat    = defaultdict(float)
        resp_renov = Counter()

        for row in rows[1:]:
            nome = row[1] if len(row) > 1 else None
            if not nome or str(nome).strip() in ("None", ""):
                continue
            # Busca colunas por palavras-chave ASCII (robusto a encoding)
            sit_col  = next((i for k, i in col.items() if "SITU" in k.upper()), 11)
            val_col  = next((i for k, i in col.items() if "VALOR" in k.upper() and "DIO" in k.upper()), 9)
            resp_col = next((i for k, i in col.items() if "RESPON" in k.upper()), 12)
            sit_raw = row[sit_col]  if len(row) > sit_col  else None
            val     = row[val_col]  if len(row) > val_col  else None
            resp    = row[resp_col] if len(row) > resp_col else None
            if not sit_raw:
                continue
            sit = str(sit_raw).strip()
            if   "Renovado"   in sit: sit = "Renovado"
            elif "Baixa"      in sit: sit = "Baixa"
            elif "Pendente"   in sit: sit = "Pendente"
            elif "Prorrogado" in sit: sit = "Prorrogado"
            else:
                continue
            sit_count[sit] += 1
            if val and not isinstance(val, str):
                sit_fat[sit] += float(val)
            if sit == "Renovado" and resp:
                resp_renov[str(resp).strip()] += 1

        renovado   = sit_count.get("Renovado", 0)
        baixa      = sit_count.get("Baixa", 0)
        taxa_renov = round(renovado / (renovado + baixa) * 100, 1) if (renovado + baixa) else 0

        resultado["renovacoes"] = {
            "situacoes":  dict(sit_count),
            "fat":        {k: round(v) for k, v in sit_fat.items()},
            "taxa_renov": taxa_renov,
            "por_resp":   sorted(
                [{"nome": k, "qtd": v} for k, v in resp_renov.items()],
                key=lambda x: -x["qtd"]
            ),
        }
    except Exception as e:
        resultado["renovacoes"] = {"erro": str(e)}

    # A.A. MAIO
    try:
        ws = wb["A.A. Maio"]
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        col = {str(h).strip(): i for i, h in enumerate(header) if h}
        total = realiz = atendim = convers = 0
        for row in rows[1:]:
            nome = row[1] if len(row) > 1 else None
            if not nome or str(nome).strip() in ("None", ""):
                continue
            total += 1
            if str(row[col.get("AULA REALIZADA?",        7)]).strip() == "Sim": realiz  += 1
            if str(row[col.get("ATENDIMENTO REALIZADO?", 8)]).strip() == "Sim": atendim += 1
            if str(row[col.get("CONVERSÃO",             11)]).strip() == "Sim": convers += 1
        resultado["aa"] = {
            "total":         total,
            "realizadas":    realiz,
            "atendimentos":  atendim,
            "conversoes":    convers,
            "taxa_presenca": round(realiz / total * 100, 1) if total else 0,
            "taxa_conv":     round(convers / realiz * 100, 1) if realiz else 0,
        }
    except Exception as e:
        resultado["aa"] = {"erro": str(e)}

    resultado["meta_dia"]   = round(META_VENDAS / 31, 2)
    resultado["meta_total"] = META_VENDAS
    resultado["atualizado"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    return resultado


@app.route("/api/data")
def api_data():
    dados = ler_planilha()
    resp = jsonify(dados)
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/")
def index():
    return send_from_directory(".", "dashboard_vendas.html")


if __name__ == "__main__":
    print("Rodando em http://localhost:" + str(PORT))
    app.run(host="0.0.0.0", port=PORT, debug=False)
