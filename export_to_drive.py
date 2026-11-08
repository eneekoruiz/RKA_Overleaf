"""
export_to_drive.py - Compiles project on Overleaf and exports the PDF directly to Google Drive.
"""
from pathlib import Path
from verify_layout import compile_overleaf_and_download

PROJECT_ID = "6acd1f92fb154e211b0d0ba7"
DRIVE_DIR = Path(
    r"G:\Mi unidad\Borrar - Dualismo filosófico - Eneko\Espiritualismo\Estudios\Euskal Herriko Unibertsitatea\Compartido Eneko\4. Kurtsoa\4. Kurtsoa 1. Lauhilekoa\RKA\rka_6\2. Proiektua"
)
LOCAL_PDF = Path("overleaf_downloaded.pdf")
DRIVE_PDF = DRIVE_DIR / "Praktika_Txostena_RKA.pdf"
COOKIE_FILE = Path(".olauth")

def main():
    if not COOKIE_FILE.exists():
        raise RuntimeError("No se encontró el archivo .olauth con la cookie de Overleaf")
    cookie_val = COOKIE_FILE.read_text(encoding="utf-8").strip().removeprefix("overleaf_session2=")
    
    print(f"[*] Solicitando compilación en Overleaf para el proyecto {PROJECT_ID}...")
    pdf_path = compile_overleaf_and_download(PROJECT_ID, cookie_val, LOCAL_PDF)
    print(f"[OK] PDF descargado de Overleaf: {pdf_path.stat().st_size} bytes")

    if DRIVE_DIR.exists():
        DRIVE_PDF.write_bytes(pdf_path.read_bytes())
        print(f"[OK] Exportado exitosamente a Google Drive: {DRIVE_PDF}")
    else:
        print(f"[WARN] No se detectó la unidad de Google Drive en {DRIVE_DIR}")

if __name__ == "__main__":
    main()
