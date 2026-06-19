from pathlib import Path
import shutil
from datetime import datetime

def fazer_backup(arq_origem, loc_destino, logger, manter=30):
    origem = Path(arq_origem)
    backup_dir = Path(loc_destino).parent / "backup_"
    backup_dir.mkdir(exist_ok=True)

    destino = backup_dir / f"{origem.stem}{datetime.now().strftime('_%Y-%m-%d_%H-%M')}{origem.suffix}"

    try:
        shutil.copy2(origem, destino)
        logger.info(f"Backup realizado com sucesso: {destino}")
    except Exception as e:
        logger.exception(f"Erro ao realizar backup: {e}")
        raise

    backups = sorted(backup_dir.glob(f"{origem.stem}_*{origem.suffix}"),
                    reverse=True)

    for arquivo in backups[manter:]:
        arquivo.unlink()