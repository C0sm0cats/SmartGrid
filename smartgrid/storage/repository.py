from __future__ import annotations

import copy
import json
import logging
import os
from pathlib import Path
import tempfile
from datetime import datetime, timezone

from smartgrid.core.models import Settings, SpaceProfile, LayoutTemplate

VERSION = 1
MAX_FILE_BYTES = 8 * 1024 * 1024
log = logging.getLogger(__name__)


def default_directory() -> Path:
    if os.name == "nt":
        return Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData/Local")) / "SmartGrid"
    return Path(os.environ.get("XDG_STATE_HOME", Path.home() / ".local/state")) / "SmartGrid"


def read_json(path: Path) -> dict:
    if path.stat().st_size > MAX_FILE_BYTES:
        raise ValueError("The file is too large (maximum 8 MB).")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("The file must contain a JSON object.")
    version = data.get("version", VERSION)
    if isinstance(version, bool) or not isinstance(version, int) or version != VERSION:
        raise ValueError(f"Unsupported file version: {version!r}")
    return data


def atomic_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
    fd, temporary = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


def durable_profile(profile: SpaceProfile) -> dict:
    from dataclasses import asdict
    data = asdict(profile)
    for assignment in data.get("assignments", []):
        if assignment:
            assignment["window_id"] = None
    return data


def durable_template(template: LayoutTemplate) -> dict:
    from dataclasses import asdict
    data = asdict(template)
    for assignment in data.get("assignments", []):
        if assignment:
            assignment["window_id"] = None
    return data


class Repository:
    def __init__(self, directory: str | Path | None = None, *, read_only=False):
        self.directory = Path(directory) if directory else default_directory()
        self.read_only = read_only
        self.errors: list[str] = []
        self.variants: dict[str, dict] = {}

    def _read(self, filename: str) -> dict:
        path = self.directory / filename
        if not path.exists():
            return {}
        try:
            return read_json(path)
        except (OSError, ValueError, TypeError) as error:
            self.errors.append(f"{filename}: {error}")
            log.exception("Unable to read %s; preserving original", path)
            self._preserve_invalid(filename)
            return {}

    def _preserve_invalid(self, filename):
        if self.read_only:
            return
        path = self.directory / filename
        if path.exists():
            import shutil
            stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
            try:
                shutil.copy2(path, path.with_name(f"{filename}.{stamp}.invalid"))
            except OSError:
                log.exception("Unable to preserve invalid configuration")

    def load_settings(self) -> Settings:
        data = self._read("settings.json").get("settings", {})
        try:
            return Settings.from_dict(data)
        except (ValueError, TypeError, KeyError) as error:
            self.errors.append(f"Invalid preferences: {error}")
            self._preserve_invalid("settings.json")
            return Settings()

    def save_settings(self, settings: Settings) -> None:
        if self.read_only:
            raise RuntimeError("Diagnostics cannot write preferences.")
        from dataclasses import asdict
        validated = Settings.from_dict(asdict(settings))
        atomic_json(self.directory / "settings.json", {"version": VERSION, "settings": asdict(validated)})

    def load_layouts(self) -> tuple[list[LayoutTemplate], dict[tuple[str, int], SpaceProfile]]:
        data = self._read("layouts.json")
        templates: list[LayoutTemplate] = []
        profiles: dict[tuple[str, int], SpaceProfile] = {}
        raw_templates, raw_profiles = data.get("templates", []), data.get("profiles", [])
        if not isinstance(raw_templates, list):
            self.errors.append("The saved layouts are invalid and were ignored.")
            raw_templates = []
        if not isinstance(raw_profiles, list):
            self.errors.append("The space profiles are invalid and were ignored.")
            raw_profiles = []
        for value in raw_templates:
            try:
                template = LayoutTemplate.from_dict(value)
                if any(existing.id == template.id for existing in templates):
                    raise ValueError("Duplicate layout ID")
                templates.append(template)
            except (ValueError, TypeError, KeyError) as error:
                self.errors.append(f"Layout ignored: {error}")
        for value in raw_profiles:
            try:
                profile = SpaceProfile.from_dict(value)
                profiles[(profile.display_id, profile.space)] = profile
            except (ValueError, TypeError, KeyError) as error:
                self.errors.append(f"Profile ignored: {error}")
        variants = data.get("variants", {})
        if isinstance(variants, dict):
            for key, value in variants.items():
                try:
                    self.variants[str(key)] = durable_profile(SpaceProfile.from_dict(value))
                except (ValueError, TypeError, KeyError):
                    self.errors.append(f"Variant ignored: {key}")
        if any("ignor" in error.casefold() or "invalid" in error.casefold() for error in self.errors):
            self._preserve_invalid("layouts.json")
        return templates, profiles

    def remember_variant(self, profile: SpaceProfile) -> None:
        key = json.dumps([profile.display_id, profile.space, profile.preset], ensure_ascii=False)
        self.variants[key] = durable_profile(profile)

    def get_variant(self, display_id: str, space: int, preset: str) -> SpaceProfile | None:
        value = self.variants.get(json.dumps([display_id, space, preset], ensure_ascii=False))
        return SpaceProfile.from_dict(copy.deepcopy(value)) if value else None

    def save_layouts(self, templates, profiles) -> None:
        if self.read_only:
            raise RuntimeError("Diagnostics cannot write profiles.")
        atomic_json(self.directory / "layouts.json", {
            "version": VERSION,
            "templates": [durable_template(t) for t in templates],
            "profiles": [durable_profile(p) for p in profiles.values()],
            "variants": copy.deepcopy(self.variants),
        })

    def load_session(self) -> dict:
        data=self._read('session.json')
        pending=data.get('pending',[]);spaces=data.get('active_spaces',{})
        if not isinstance(pending,list) or not isinstance(spaces,dict):
            self.errors.append('Invalid runtime session');self._preserve_invalid('session.json');return {}
        clean=[]
        import math
        for entry in pending:
            if not isinstance(entry,dict): continue
            display,space,index,app,created=(entry.get(k) for k in ('display_id','space','index','app_id','created_utc'))
            if (isinstance(display,str) and display and type(space) is int and space in range(3)
                    and type(index) is int and 0<=index<10000 and isinstance(app,str) and app
                    and type(created) in (int,float) and math.isfinite(created) and created>0):
                clean.append(dict(display_id=display,space=space,index=index,app_id=app,created_utc=created))
        return {'pending':clean,'active_spaces':{k:v for k,v in spaces.items() if isinstance(k,str) and type(v) is int and v in range(3)}}

    def save_session(self, pending, active_spaces):
        if self.read_only: return
        atomic_json(self.directory/'session.json',{'version':VERSION,'pending':pending,'active_spaces':active_spaces})
