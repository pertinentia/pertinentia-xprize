import sys
sys.path.insert(0, '.')
import time

print("Precalentando caché de plantillas/ ...")
t0 = time.time()
import ai_helper
docs = ai_helper.load_all_reference_docs()
t1 = time.time()
print(f"✅ {len(docs)} archivos procesados en {t1-t0:.1f}s")
for fname, text in docs.items():
    estado = "OK" if text and not text.startswith("[Error") else "⚠️ VACÍO O ERROR"
    print(f"  {fname}: {len(text or '')} caracteres — {estado}")
