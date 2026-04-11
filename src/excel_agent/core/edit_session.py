"""Edit session abstraction for unified "edit target" semantics.

Provides a single entry point for all mutating operations:
- Resolves edit target (input vs output)
- Handles file copying when needed
- Provides locked ExcelAgent context
- Eliminates double-save issues

Usage:
    with EditSession.prepare(input_path, output_path) as session:
        # session.path is the file being modified
        # session.agent provides locked ExcelAgent
        # session.is_inplace tells if input == output
        wb = session.workbook
        # ... make modifications ...
        # Auto-saved to correct path on exit

The EditSession ensures:
1. Every tool knows exactly which file it is modifying
2. File locking happens consistently via ExcelAgent
3. Double-save is impossible (only ExcelAgent.__exit__ saves)
4. Macro preservation is consistent (via ExcelAgent's keep_vba)
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import TYPE_CHECKING, Optional

from excel_agent.core.agent import ExcelAgent
from excel_agent.utils.exceptions import ValidationError

if TYPE_CHECKING:
    from openpyxl import Workbook

logger = logging.getLogger(__name__)


class EditSession:
    """Manages edit session lifecycle with unified target semantics.

    This class eliminates the double-save bug by ensuring only ExcelAgent
    ever calls wb.save(). The edit target is resolved at preparation time:
    - In-place edits: edit the input file directly
    - Copy edits: copy input to output, edit the output

    Pattern:
        session = EditSession.prepare(input_path, output_path)
        with session:
            # session.path is the file being modified
            # session.agent provides locked ExcelAgent
            # session.is_inplace tells if input == output
            pass
        # Auto-saved to correct path on exit
    """

    def __init__(
        self,
        edit_path: Path,
        is_inplace: bool,
        input_path: Optional[Path] = None,
    ) -> None:
        """Initialize EditSession (use prepare() factory method).

        Args:
            edit_path: The actual file path being edited
            is_inplace: Whether this is an in-place edit
            input_path: Original input path (for reporting)
        """
        self.edit_path = Path(edit_path).resolve()
        self.is_inplace = is_inplace
        self.input_path = Path(input_path).resolve() if input_path else self.edit_path
        self._agent: Optional[ExcelAgent] = None

    @classmethod
    def prepare(
        cls,
        input_path: Path | str,
        output_path: Optional[Path | str] = None,
        *,
        force_inplace: bool = True,  # Default True for backward compatibility
        create_parents: bool = True,
    ) -> "EditSession":
        """Prepare edit session with unified semantics.

        Args:
            input_path: Source file path (must exist)
            output_path: Target output path (None or same as input = inplace)
            force_inplace: Allow editing input directly (default True)
            create_parents: Create output parent directories if needed

        Returns:
            EditSession configured for the resolved edit target

        Raises:
            FileNotFoundError: If input file doesn't exist
        """
        input_path = Path(input_path).resolve()

        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")

        if not input_path.is_file():
            raise ValueError(f"Input path is not a file: {input_path}")

        # If no output specified, or output same as input, do in-place edit
        if output_path is None:
            output_path = input_path

        output_path = Path(output_path).resolve()

        if output_path == input_path:
            # In-place edit (default behavior)
            logger.debug("EditSession: In-place edit of %s", input_path)
            return cls(edit_path=input_path, is_inplace=True, input_path=input_path)

        # Different output: copy input to output, edit output
        if create_parents:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        elif not output_path.parent.exists():
            raise FileNotFoundError(f"Output directory does not exist: {output_path.parent}")

        # Copy input to output before editing
        logger.debug("EditSession: Copy %s -> %s for editing", input_path, output_path)
        shutil.copy2(input_path, output_path)

        return cls(
            edit_path=output_path,
            is_inplace=False,
            input_path=input_path,
        )

    def __enter__(self) -> "EditSession":
        """Enter locked ExcelAgent context.

        Returns:
            self for context manager use

        Raises:
            RuntimeError: If already entered
        """
        if self._agent is not None:
            raise RuntimeError("EditSession already entered")

        logger.debug("EditSession entering: %s", self.edit_path)
        self._agent = ExcelAgent(self.edit_path, mode="rw").__enter__()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Exit and save via ExcelAgent.

        ExcelAgent handles the save to self.edit_path.
        This is the ONLY place wb.save() is called.
        """
        if self._agent is not None:
            logger.debug(
                "EditSession exiting: %s (exception=%s)",
                self.edit_path,
                exc_type.__name__ if exc_type else "None",
            )
            self._agent.__exit__(exc_type, exc_val, exc_tb)
            self._agent = None

    @property
    def workbook(self) -> "Workbook":
        """Access workbook through agent.

        Returns:
            The openpyxl Workbook object

        Raises:
            RuntimeError: If EditSession not entered
        """
        if self._agent is None:
            raise RuntimeError(
                "EditSession.workbook accessed outside context manager. "
                "Use 'with EditSession.prepare(...) as session:'"
            )
        return self._agent.workbook

    @property
    def agent(self) -> ExcelAgent:
        """Access underlying agent for version hash, etc.

        Returns:
            The ExcelAgent instance

        Raises:
            RuntimeError: If EditSession not entered
        """
        if self._agent is None:
            raise RuntimeError(
                "EditSession.agent accessed outside context manager. "
                "Use 'with EditSession.prepare(...) as session:'"
            )
        return self._agent

    @property
    def version_hash(self) -> str:
        """Get geometry hash for version reporting.

        Returns:
            The geometry hash from the underlying agent
        """
        if self._agent is None:
            raise RuntimeError("EditSession.version_hash accessed outside context manager")
        return self._agent.version_hash

    @property
    def file_hash(self) -> str:
        """Get file hash for version reporting.

        Returns:
            The file hash from the underlying agent
        """
        if self._agent is None:
            raise RuntimeError("EditSession.file_hash accessed outside context manager")
        return self._agent.file_hash


def prepare_edit_target(
    input_path: str | Path,
    output_path: Optional[str | Path] = None,
    *,
    create_parents: bool = True,
    force_inplace: bool = False,
) -> Path:
    """Legacy compatibility: resolve edit target path.

    Returns the path that should be opened for editing.
    This is the file that EditSession will actually modify.

    Args:
        input_path: Source file path
        output_path: Target output path (None = inplace)
        create_parents: Create output parent directories
        force_inplace: Allow editing input directly

    Returns:
        The resolved edit target path

    Examples:
        >>> prepare_edit_target("test.xlsx")
        PosixPath('/path/to/test.xlsx')

        >>> prepare_edit_target("input.xlsx", "output.xlsx")
        PosixPath('/path/to/output.xlsx')  # File copied here
    """
    inp = Path(input_path).resolve()
    out = Path(output_path).resolve() if output_path else None

    session = EditSession.prepare(
        inp,
        out,
        create_parents=create_parents,
        force_inplace=force_inplace,
    )
    return session.edit_path
