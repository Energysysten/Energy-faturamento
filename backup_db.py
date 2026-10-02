"""
Backup do banco faturamento.db.

- Usa a API de backup do SQLite: cópia consistente mesmo com o sistema rodando.
- Verifica a integridade da cópia antes de guardar.
- Compacta em .zip (abre direto no Windows) e grava em cada pasta de BACKUP_DIRS.
- Retenção: mantém os últimos 30 backups + o primeiro de cada mês por 12 meses.

Rodar no VM: python backup_db.py
Variáveis opcionais:
  FAT_DB        caminho do banco (padrão: faturamento.db ao lado deste script)
  BACKUP_DIRS   pastas de destino separadas por ";"
  BACKUP_KEEP   quantidade de backups recentes a manter (padrão 30)
  BACKUP_MESES  meses de backup mensal a manter (padrão 12)
"""
import os, sys, sqlite3, zipfile, tempfile, shutil
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("FAT_DB", BASE / "faturamento.db"))
DEFAULT_DIRS = [
    str(BASE / "backups"),
    r"\\Zion2\financeiro\13 - FOLHAS_SISTEMAFATURAMENTO\BACKUP_BANCO",
]
BACKUP_DIRS = [d.strip() for d in os.environ.get("BACKUP_DIRS", ";".join(DEFAULT_DIRS)).split(";") if d.strip()]
KEEP = int(os.environ.get("BACKUP_KEEP", 30))
MESES = int(os.environ.get("BACKUP_MESES", 12))
PREFIX = "faturamento_"
LOG_PATH = BASE / "backups" / "backup.log"


def log(msg):
    linha = f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    print(linha, flush=True)
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except OSError:
        pass


def gerar_zip(tmpdir):
    """Cria snapshot consistente, valida e compacta. Retorna o caminho do .zip."""
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    snap = Path(tmpdir) / "faturamento.db"
    src = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    dst = sqlite3.connect(snap)
    try:
        src.backup(dst)
        res = dst.execute("PRAGMA integrity_check").fetchone()[0]
        if res != "ok":
            raise RuntimeError(f"integrity_check falhou: {res}")
        tabelas = dst.execute("SELECT count(*) FROM sqlite_master WHERE type='table'").fetchone()[0]
        if tabelas == 0:
            raise RuntimeError("cópia sem nenhuma tabela")
    finally:
        dst.close()
        src.close()
    zpath = Path(tmpdir) / f"{PREFIX}{stamp}.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(snap, "faturamento.db")
    return zpath, tabelas


def aplicar_retencao(pasta):
    arquivos = sorted(pasta.glob(f"{PREFIX}*.zip"), reverse=True)  # mais novo primeiro
    manter = set(arquivos[:KEEP])
    meses_vistos = {}
    for a in sorted(arquivos):  # mais antigo primeiro → pega o 1º de cada mês
        mes = a.name[len(PREFIX):len(PREFIX) + 7]  # AAAA-MM
        meses_vistos.setdefault(mes, a)
    for mes in sorted(meses_vistos, reverse=True)[:MESES]:
        manter.add(meses_vistos[mes])
    removidos = 0
    for a in arquivos:
        if a not in manter:
            a.unlink()
            removidos += 1
    return removidos


def main():
    if not DB_PATH.exists():
        log(f"ERRO banco não encontrado: {DB_PATH}")
        return 1
    with tempfile.TemporaryDirectory() as tmp:
        try:
            zpath, tabelas = gerar_zip(tmp)
        except Exception as e:
            log(f"ERRO ao gerar backup: {e}")
            return 1
        tam = zpath.stat().st_size / 1024
        falhas = 0
        for d in BACKUP_DIRS:
            pasta = Path(d)
            try:
                pasta.mkdir(parents=True, exist_ok=True)
                shutil.copy2(zpath, pasta / zpath.name)
                rem = aplicar_retencao(pasta)
                log(f"OK {pasta / zpath.name} ({tam:.0f} KB, {tabelas} tabelas, {rem} antigos removidos)")
            except Exception as e:
                falhas += 1
                log(f"ERRO ao gravar em {pasta}: {e}")
    if falhas == len(BACKUP_DIRS):
        return 1
    return 2 if falhas else 0


if __name__ == "__main__":
    sys.exit(main())
