"""
Gestión de escenarios: guardar, cargar y eliminar configuraciones como JSON.
"""
import os
import json
import re
from typing import List, Optional

from model import ProjectionParams

SCENARIOS_DIR = os.path.join(os.path.dirname(__file__), "scenarios")


def _ensure_dir() -> None:
    os.makedirs(SCENARIOS_DIR, exist_ok=True)


def _safe_name(name: str) -> str:
    """Convierte el nombre a nombre de archivo seguro."""
    safe = re.sub(r"[^\w\s\-]", "", name, flags=re.UNICODE).strip()
    safe = re.sub(r"\s+", "_", safe)
    return safe or "escenario"


def list_scenarios() -> List[str]:
    """Lista nombres de escenarios guardados (sin extensión)."""
    _ensure_dir()
    return sorted(
        f[:-5] for f in os.listdir(SCENARIOS_DIR) if f.endswith(".json")
    )


def save_scenario(name: str, params: ProjectionParams) -> bool:
    """Guarda escenario. Retorna True si exitoso."""
    _ensure_dir()
    path = os.path.join(SCENARIOS_DIR, f"{_safe_name(name)}.json")
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(params.to_dict(), fh, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def load_scenario(name: str) -> Optional[ProjectionParams]:
    """Carga escenario por nombre. Retorna None si no existe o hay error."""
    _ensure_dir()
    path = os.path.join(SCENARIOS_DIR, f"{_safe_name(name)}.json")
    if not os.path.exists(path):
        # También intentar con el nombre exacto (ya viene con safe_name)
        path = os.path.join(SCENARIOS_DIR, f"{name}.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return ProjectionParams.from_dict(data)
    except (OSError, KeyError, TypeError, json.JSONDecodeError):
        return None


def delete_scenario(name: str) -> bool:
    """Elimina escenario. Retorna True si fue eliminado."""
    _ensure_dir()
    path = os.path.join(SCENARIOS_DIR, f"{_safe_name(name)}.json")
    if not os.path.exists(path):
        path = os.path.join(SCENARIOS_DIR, f"{name}.json")
    if os.path.exists(path):
        os.remove(path)
        return True
    return False
