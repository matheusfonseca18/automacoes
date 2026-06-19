from pathlib import Path
import shutil
from datetime import datetime

planilha_baixada_path =r"C:\Users\matheus.pinto\Desktop\Atividades\Almap Africa\almap africa.csv"
planilha_final_path =r"C:\Users\matheus.pinto\Desktop\Atividades\Indicador class\Indicador - classificação 2026.xlsx"

arquivo = Path(planilha_final_path)

lock_file = arquivo.parent / f"~${arquivo.name}"

if lock_file.exists():
    print("Planilha está aberta por alguém.")
    exit()
else:
    print("Planilha não está aberta por ninguém.")

print("Seguiu")



# origem = Path(planilha_final_path)
# backup_dir = Path(planilha_baixada_path).parent / "backup_"
# backup_dir.mkdir(exist_ok=True)

# destino = backup_dir / f"{origem.stem}{datetime.now().strftime('_%Y-%m-%d_%H-%M')}{origem.suffix}"

# shutil.copy2(origem, destino)

# backups = sorted(backup_dir.glob(f"{origem.stem}_*{origem.suffix}"),
#                  reverse=True)

# for arquivo in backups[30:]:
#     arquivo.unlink()