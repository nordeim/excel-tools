"""Utility modules for excel-agent-tools."""

__all__ = [
    "ExitCode",
    "build_response",
    "print_json",
]


def __getattr__(name: str) -> object:
    if name == "ExitCode":
        from excel_agent.utils.exit_codes import ExitCode

        return ExitCode
    if name in ("build_response", "print_json"):
        import excel_agent.utils.json_io as json_io

        return getattr(json_io, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
