"""Observable outcomes of non-atomic operations on third-party windows."""
from dataclasses import dataclass, field

from .models import Rect


@dataclass
class ApplyResult:
    placed: list[int] = field(default_factory=list)
    hidden: list[int] = field(default_factory=list)
    launched: list[str] = field(default_factory=list)
    unavailable: list[str] = field(default_factory=list)
    failures: dict[int, str] = field(default_factory=dict)

    @property
    def success(self):
        return not self.failures and not self.unavailable

    @property
    def summary(self):
        parts = [f"{len(self.placed)} windows placed"]
        if self.hidden:
            parts.append(f"{len(self.hidden)} minimized")
        if self.launched:
            parts.append(f"{len(self.launched)} apps opened")
        if self.unavailable:
            parts.append(f"{len(self.unavailable)} unavailable")
        if self.failures:
            parts.append(f"{len(self.failures)} placements failed")
        return ", ".join(parts)
