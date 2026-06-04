"""
abrir_dashboard.py — Launcher local do Dashboard O&M Parnaíba
──────────────────────────────────────────────────────────────
Uso:
    python abrir_dashboard.py          # abre com D-1 (padrão)
    python abrir_dashboard.py --agendar  # cria tarefa Windows 00:01

Fluxo:
  1. Verifica se o CSV de PDP D-1 existe; se não, executa pdp_download.py
  2. Sobe servidor HTTP local na porta 8765 (serve pasta dev/)
  3. Expõe endpoint GET /api/pdp?data=YYYY-MM-DD  → JSON { "HH:MM": { "UG": MW } }
  4. Abre o browser em http://localhost:8765/dashboard_intradiario_parnaiba.html
"""

import argparse, csv, http.server, io, json, os, subprocess, sys, threading, webbrowser
_pdp_downloading = set()   # datas com download em andamento
from datetime    import datetime, timedelta
from pathlib     import Path
from urllib.parse import urlparse, parse_qs

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

DEV_DIR = Path(__file__).parent
PORTA   = 8765

# ─── Carregar PDP do CSV local ────────────────────────────────────────────────

def carregar_pdp(data_iso: str) -> dict:
    """
    Lê PDP_ENEVA_N_{data_iso}.csv e retorna:
    { "HH:MM": { "UG_ID": mw_por_ug, ... }, ... }
    """
    arq = DEV_DIR / f"PDP_ENEVA_N_{data_iso}.csv"
    if not arq.exists():
        return {}

    mapa: dict = {}
    with open(str(arq), encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f):
            hora = row.get("Hora", "").strip()
            if not hora: continue
            # normalizar "24:00" → "00:00" (alguns sistemas usam 24h)
            if hora == "24:00": hora = "00:00"
            ugs_str = row.get("UGs_Dashboard", "")
            ugs = [u.strip() for u in ugs_str.split(",") if u.strip()]
            n   = len(ugs)
            if n == 0: continue
            try:
                mw_total = float(row.get("MW_Programado", "0").replace(",", "."))
            except ValueError:
                continue
            mw_por = round(mw_total / n, 1)
            if hora not in mapa:
                mapa[hora] = {}
            for ug in ugs:
                mapa[hora][ug] = mw_por

    return mapa

# ─── Carregar Way2 do CSV local ───────────────────────────────────────────────

def carregar_way2(data_iso: str, res: int) -> dict:
    """
    Lê Way2_UG_{res}min_{data_iso}.csv e retorna:
    { "date": ..., "resolution": ..., "labels": ["HH:MM", ...], "data": { "unitId": [...] } }
    """
    arq = DEV_DIR / f"Way2_UG_{res}min_{data_iso}.csv"
    if not arq.exists():
        return {}

    labels: list = []
    data:   dict = {}

    with open(str(arq), encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        unit_cols = [c for c in (reader.fieldnames or []) if c != "DataHora"]
        for col in unit_cols:
            data[col] = []
        for row in reader:
            dt_str = row.get("DataHora", "")
            # label = HH:MM (posição 11-16 de "YYYY-MM-DD HH:MM:SS")
            label  = dt_str[11:16] if len(dt_str) >= 16 else dt_str
            labels.append(label)
            for col in unit_cols:
                v = row.get(col, "")
                try:
                    data[col].append(float(v) if v else None)
                except ValueError:
                    data[col].append(None)

    return {"date": data_iso, "resolution": res, "labels": labels, "data": data}


# ─── Handler HTTP ─────────────────────────────────────────────────────────────

class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(DEV_DIR), **kwargs)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/pdp":
            self._serve_pdp(parse_qs(parsed.query))
        elif parsed.path == "/api/way2":
            self._serve_way2(parse_qs(parsed.query))
        elif parsed.path == "/api/status":
            self._serve_status()
        elif parsed.path == "/api/coletar":
            self._serve_coletar(parse_qs(parsed.query))
        else:
            super().do_GET()

    def _serve_pdp(self, params):
        data = (params.get("data") or [None])[0]
        if not data:
            data = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

        arq = DEV_DIR / f"PDP_ENEVA_N_{data}.csv"
        if not arq.exists() and data not in _pdp_downloading:
            _pdp_downloading.add(data)
            def _baixar():
                try:
                    subprocess.run(
                        [sys.executable, str(DEV_DIR / "pdp_download.py"), "--data", data],
                        cwd=str(DEV_DIR),
                    )
                finally:
                    _pdp_downloading.discard(data)
            threading.Thread(target=_baixar, daemon=True).start()
            print(f"[pdp] Download iniciado em background para {data}")

        mapa = carregar_pdp(data)
        self._json(mapa)

    def _serve_coletar(self, params):
        data = (params.get("data") or [None])[0]
        if not data:
            self._json({"ok": False, "erro": "Parâmetro data ausente"})
            return
        print(f"[api/coletar] Iniciando coleta Way2 para {data}...")
        ret = subprocess.run(
            [sys.executable, str(DEV_DIR / "way2_coleta.py"), "--data", data],
            cwd=str(DEV_DIR),
        )
        if ret.returncode == 0:
            print(f"[api/coletar] Concluído: {data}")
            self._json({"ok": True})
        else:
            print(f"[api/coletar] Falha: {data}")
            self._json({"ok": False, "erro": "Coleta Way2 falhou (data não suportada ou erro de rede)"})

    def _serve_status(self):
        hoje = datetime.now().date()
        faltando = []
        for i in range(1, 8):
            d = hoje - timedelta(days=i)
            data_iso = d.strftime("%Y-%m-%d")
            arq30 = DEV_DIR / f"Way2_UG_30min_{data_iso}.csv"
            arq5  = DEV_DIR / f"Way2_UG_5min_{data_iso}.csv"
            if not (arq30.exists() and arq5.exists()):
                faltando.append(data_iso)
        self._json({"faltando": faltando})

    def _serve_way2(self, params):
        data = (params.get("data") or [None])[0]
        res  = int((params.get("res") or ["30"])[0])
        if not data:
            data = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

        resultado = carregar_way2(data, res)
        self._json(resultado)

    def _json(self, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type",   "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def end_headers(self):
        # Impede cache do browser para arquivos HTML
        if self.path.split("?")[0].endswith(".html"):
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, fmt, *args):
        if args and str(args[1]) not in ("200", "304"):
            print(f"[servidor] {fmt % args}")

# ─── Garantir CSV disponível ──────────────────────────────────────────────────

def garantir_way2(data_iso: str, forcar: bool = False) -> bool:
    arq5  = DEV_DIR / f"Way2_UG_5min_{data_iso}.csv"
    arq30 = DEV_DIR / f"Way2_UG_30min_{data_iso}.csv"
    if arq5.exists() and arq30.exists():
        if forcar:
            # conta slots do 30min (48 = dia completo); re-coleta só se incompleto
            with open(str(arq30), encoding="utf-8") as _f:
                slots = sum(1 for _ in _f) - 1  # desconta header
            if slots >= 48:
                print(f"[v] Way2 D-1 já completo ({arq30.name}, {slots} slots).")
                return True
            print(f"[i] Way2 D-1 incompleto ({slots} slots) — re-coletando ({data_iso})...")
            arq5.unlink()
            arq30.unlink()
        else:
            print(f"[v] Way2 encontrado: {arq5.name}")
            return True
    else:
        print(f"[i] Way2 CSV não encontrado — executando way2_coleta.py ({data_iso})...")
    ret = subprocess.run(
        [sys.executable, str(DEV_DIR / "way2_coleta.py"), "--data", data_iso],
        cwd=str(DEV_DIR),
    )
    if ret.returncode == 0:
        print("[v] Way2 coletado.")
        return True
    print("[!] Coleta Way2 falhou. Dashboard abrirá sem dados de geração em tempo real.")
    return False


def garantir_pdp(data_iso: str) -> bool:
    arq = DEV_DIR / f"PDP_ENEVA_N_{data_iso}.csv"
    if arq.exists():
        print(f"[v] PDP encontrado: {arq.name}")
        return True
    print(f"[i] CSV não encontrado — baixando D-1 ({data_iso})...")
    ret = subprocess.run(
        [sys.executable, str(DEV_DIR / "pdp_download.py"), "--data", data_iso],
        cwd=str(DEV_DIR),
    )
    if ret.returncode == 0:
        print("[v] Download concluído.")
        return True
    print("[!] Download falhou. Dashboard abrirá sem PDP local.")
    return False

# ─── Agendador Windows ────────────────────────────────────────────────────────

def criar_tarefa_agendada():
    script = Path(__file__).resolve()
    python = sys.executable
    cmd = (
        f'schtasks /create /tn "Dashboard_PDP_Parnaiba" '
        f'/tr "\\"{python}\\" \\"{script}\\"" '
        f'/sc daily /st 00:01 /ru "{os.environ.get("USERNAME","")}" /f'
    )
    print(f"Criando tarefa agendada: {cmd}")
    ret = os.system(cmd)
    if ret == 0:
        print("[v] Tarefa criada: Dashboard_PDP_Parnaiba — diário 00:01")
        print("    Verifique em: Agendador de Tarefas → Biblioteca do Agendador")
    else:
        print("[X] Falha ao criar tarefa (execute como Administrador)")

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Launcher local do Dashboard Parnaíba")
    ap.add_argument("--agendar", action="store_true",
                    help="Criar tarefa no Agendador de Tarefas Windows (00:01)")
    args = ap.parse_args()

    if args.agendar:
        criar_tarefa_agendada()
        return

    d0 = datetime.now().strftime("%Y-%m-%d")
    d1 = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    garantir_way2(d1, forcar=True)
    garantir_pdp(d1)
    garantir_way2(d0)
    # PDP D-0: download em background via /api/pdp (não bloqueia inicialização)

    # Verifica se a porta já está em uso (ex: servidor anterior ainda ativo)
    import socket as _sock
    _s = _sock.socket(_sock.AF_INET, _sock.SOCK_STREAM)
    porta_livre = _s.connect_ex(("localhost", PORTA)) != 0
    _s.close()

    if not porta_livre:
        print(f"[i] Porta {PORTA} já em uso — coleta concluída, servidor não iniciado.")
        return

    # Servidor HTTP local
    server = http.server.HTTPServer(("localhost", PORTA), DashboardHandler)
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    print(f"[v] Servidor: http://localhost:{PORTA}")

    url = f"http://localhost:{PORTA}/dashboard_intradiario_parnaiba.html"
    try:
        chrome = webbrowser.get("chrome")
        chrome.open(url)
    except webbrowser.Error:
        # Chrome não registrado no webbrowser — tenta caminhos comuns do Windows
        import subprocess as _sp
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            rf"C:\Users\{os.environ.get('USERNAME','')}\AppData\Local\Google\Chrome\Application\chrome.exe",
        ]
        launched = False
        for p in chrome_paths:
            if Path(p).exists():
                _sp.Popen([p, url])
                launched = True
                break
        if not launched:
            webbrowser.open(url)  # fallback para browser padrão
    print(f"[v] Dashboard: {url}")
    print("    Pressione Ctrl+C para encerrar.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()
        print("\n[i] Servidor encerrado.")


if __name__ == "__main__":
    main()
