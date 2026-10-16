# ═══ HÜCRE 5 — KURULUM (yapay zekâ kütüphanesi, hücre 6 ve 8 için) ═══
import subprocess, sys, importlib
subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', 'anthropic'], check=False)
importlib.invalidate_caches()
import anthropic
print(f'✅ anthropic kuruldu (sürüm {anthropic.__version__})')
