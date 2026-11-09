import os
import json
import io
from datetime import datetime

UPLOADS_DIR = os.path.join(os.path.dirname(__file__), "uploads")
METADATA_FILE = os.path.join(UPLOADS_DIR, "metadata.json")

def ensure_uploads_dir():
    """Ensure the uploads folder exists."""
    if not os.path.exists(UPLOADS_DIR):
        os.makedirs(UPLOADS_DIR, exist_ok=True)

def load_metadata():
    """Load metadata dictionary from JSON file."""
    ensure_uploads_dir()
    if os.path.exists(METADATA_FILE):
        try:
            with open(METADATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_metadata(data):
    """Save metadata dictionary to JSON file."""
    ensure_uploads_dir()
    with open(METADATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_file(uploaded_file, file_key):
    """Save an uploaded file to disk and record metadata."""
    ensure_uploads_dir()
    target_path = os.path.join(UPLOADS_DIR, f"{file_key}.xlsx")
    
    file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, 'getvalue') else uploaded_file.read()
    with open(target_path, "wb") as f:
        f.write(file_bytes)
        
    meta = load_metadata()
    size_mb = len(file_bytes) / (1024 * 1024)
    size_str = f"{size_mb:.2f} MB" if size_mb >= 1 else f"{len(file_bytes)/1024:.1f} KB"
    
    meta[file_key] = {
        "original_name": getattr(uploaded_file, 'name', f"{file_key}.xlsx"),
        "last_updated": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "size": size_str,
        "path": target_path
    }
    save_metadata(meta)
    return target_path

def get_file_info(file_key):
    """Get status and info of saved file."""
    target_path = os.path.join(UPLOADS_DIR, f"{file_key}.xlsx")
    if not os.path.exists(target_path):
        return None
    meta = load_metadata()
    info = meta.get(file_key, {})
    return {
        "exists": True,
        "path": target_path,
        "original_name": info.get("original_name", f"{file_key}.xlsx"),
        "last_updated": info.get("last_updated", "Desconhecido"),
        "size": info.get("size", "")
    }

def delete_file(file_key):
    """Delete a saved file from disk."""
    target_path = os.path.join(UPLOADS_DIR, f"{file_key}.xlsx")
    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except Exception:
            pass
    meta = load_metadata()
    if file_key in meta:
        del meta[file_key]
        save_metadata(meta)

def get_effective_file(uploaded_file, file_key):
    """
    Returns filepath for processing.
    If uploaded_file is provided, saves it first and returns filepath.
    Otherwise, returns path to saved file if it exists.
    """
    if uploaded_file is not None:
        return save_file(uploaded_file, file_key)
    
    info = get_file_info(file_key)
    if info and os.path.exists(info["path"]):
        return info["path"]
    
    return None

def get_effective_cadastro_excesso(uploaded_exclusive=None):
    """
    Returns tuple of (file_path, source_type, info_dict).
    source_type is 'exclusive', 'shared', or 'none'.
    Prioritizes exclusive file 'ee_cadastro' if uploaded or already saved on disk.
    Falls back to shared 'planilha_1' from Ponto de Pedido if available.
    """
    if uploaded_exclusive is not None:
        save_file(uploaded_exclusive, "ee_cadastro")
        info = get_file_info("ee_cadastro")
        return info["path"], "exclusive", info

    info_ee = get_file_info("ee_cadastro")
    if info_ee and os.path.exists(info_ee["path"]):
        return info_ee["path"], "exclusive", info_ee

    info_shared = get_file_info("planilha_1")
    if info_shared and os.path.exists(info_shared["path"]):
        return info_shared["path"], "shared", info_shared

    return None, "none", None

