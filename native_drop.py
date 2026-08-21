from __future__ import annotations

from pathlib import Path
from typing import Callable, Optional

from tkinterdnd2 import DND_FILES

DropCallback = Callable[[list[Path], int, int], None]


class NativeFileDrop:
    """Safe file-drop adapter built on Tk's drag-and-drop event system."""

    def __init__(self, root, callback: DropCallback) -> None:
        self.root = root
        self.callback = callback
        self.supported = False
        self.error: Optional[str] = None
        self._binding_id = None

        try:
            self.root.drop_target_register(DND_FILES)
            self._binding_id = self.root.dnd_bind("<<Drop>>", self._on_drop)
            self.supported = True
        except Exception as exc:
            self.error = str(exc)

    def _on_drop(self, event) -> str:
        try:
            paths = [Path(value) for value in self.root.tk.splitlist(event.data) if value]
            if paths:
                self.callback(paths, int(getattr(event, "x", 0)), int(getattr(event, "y", 0)))
        except Exception as exc:
            # Do not allow a malformed external drop payload to terminate Tk.
            self.error = str(exc)
        return getattr(event, "action", "copy")

    def close(self) -> None:
        if not self.supported:
            return
        try:
            if self._binding_id:
                self.root.dnd_unbind("<<Drop>>", self._binding_id)
            self.root.drop_target_unregister()
        except Exception:
            pass
        finally:
            self.supported = False
