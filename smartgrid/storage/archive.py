from __future__ import annotations

import copy
from pathlib import Path
import uuid

from smartgrid.core.models import LayoutTemplate, SpaceProfile
from .repository import VERSION, atomic_json, read_json, durable_profile, durable_template


def export_archive(path, templates, profiles, variants=None) -> None:
    atomic_json(Path(path), {
        "format": "smartgrid-layouts",
        "version": VERSION,
        "templates": [durable_template(t) for t in templates],
        "profiles": [durable_profile(p) for p in profiles.values()],
        "variants": copy.deepcopy(variants or {}),
    })


def import_archive(path, existing_templates, existing_profiles, existing_variants=None) -> dict:
    data = read_json(Path(path))
    if data.get("format") != "smartgrid-layouts":
        raise ValueError("This is not a supported SmartGrid layout archive.")
    if not isinstance(data.get("templates", []), list) or not isinstance(data.get("profiles", []), list):
        raise ValueError("The archive contains invalid layouts or profiles.")
    templates = copy.deepcopy(existing_templates)
    profiles = copy.deepcopy(existing_profiles)
    variants = copy.deepcopy(existing_variants or {})
    names = {t.name.casefold() for t in templates}
    ids = {t.id for t in templates}
    added = renamed = kept = added_profiles = 0
    # Parse everything before returning a mutation: malformed archives are atomic failures.
    parsed_templates = [LayoutTemplate.from_dict(t) for t in data.get("templates", [])]
    parsed_profiles = [SpaceProfile.from_dict(p) for p in data.get("profiles", [])]
    for value in parsed_templates + parsed_profiles:
        for assignment in value.assignments:
            if assignment:
                assignment.window_id = None
    raw_variants = data.get("variants", {})
    if not isinstance(raw_variants, dict):
        raise ValueError("The archive contains invalid layouts or profiles.")
    parsed_variants = {str(k): durable_profile(SpaceProfile.from_dict(v)) for k, v in raw_variants.items()}
    for template in parsed_templates:
        name = template.name
        index = 1
        while name.casefold() in names:
            suffix = " (imported)" if index == 1 else f" (imported {index})"
            name = template.name[:max(1, 60 - len(suffix))] + suffix
            index += 1
        if name != template.name:
            renamed += 1
        template.name = name
        if template.id in ids:
            template.id = uuid.uuid4().hex
        names.add(template.name.casefold())
        ids.add(template.id)
        templates.append(template)
        added += 1
    for profile in parsed_profiles:
        key = (profile.display_id, profile.space)
        if key in profiles:
            kept += 1
        else:
            profiles[key] = profile
            added_profiles += 1
    for key, value in parsed_variants.items():
        variants.setdefault(key, value)
    return {"templates": templates, "profiles": profiles, "variants": variants,
            "added": added, "renamed": renamed, "profiles_kept": kept, "profiles_added": added_profiles}
