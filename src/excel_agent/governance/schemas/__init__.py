"""
JSON Schema validation utilities for excel-agent-tools.

This module provides schema loading and validation for tool inputs.
Schemas are cached in memory for performance.
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator
from jsonschema import ValidationError as JsonSchemaValidationError

from excel_agent.utils.exceptions import ValidationError

logger = logging.getLogger(__name__)

# In-memory cache for loaded schemas
_SCHEMA_CACHE: dict[str, dict[str, Any]] = {}

# Schema directory
_SCHEMA_DIR = Path(__file__).parent


def load_schema(schema_name: str) -> dict[str, Any]:
    """Load a JSON schema by name from the schemas directory.

    Schemas are cached in memory after first load for performance.

    Args:
        schema_name: Name of the schema file (e.g., "range_input.schema.json").

    Returns:
        The parsed schema as a dictionary.

    Raises:
        ValidationError: If the schema file doesn't exist or is invalid JSON.
    """
    if schema_name in _SCHEMA_CACHE:
        return _SCHEMA_CACHE[schema_name]

    schema_path = _SCHEMA_DIR / schema_name
    if not schema_path.exists():
        raise ValidationError(
            f"Schema not found: {schema_name}",
            details={"schema_dir": str(_SCHEMA_DIR), "requested": schema_name},
        )

    try:
        with open(schema_path, encoding="utf-8") as f:
            schema: dict[str, Any] = json.load(f)
    except json.JSONDecodeError as e:
        raise ValidationError(
            f"Invalid JSON in schema {schema_name}: {e}",
            details={"schema": schema_name, "error": str(e)},
        ) from e

    _SCHEMA_CACHE[schema_name] = schema
    logger.debug("Loaded and cached schema: %s", schema_name)
    return schema


def validate_against_schema(schema_name: str, data: dict[str, Any]) -> None:
    """Validate data against a named schema.

    Args:
        schema_name: Name of the schema file (e.g., "range_input.schema.json").
        data: Data to validate.

    Raises:
        ValidationError: If the data doesn't conform to the schema.
    """
    schema = load_schema(schema_name)

    try:
        # Draft202012Validator is the latest supported by jsonschema
        Draft202012Validator(schema).validate(data)
    except JsonSchemaValidationError as e:
        raise ValidationError(
            f"Schema validation failed: {e.message}",
            details={
                "schema": schema_name,
                "path": list(e.path),
                "validator": e.validator,
            },
        ) from e

    logger.debug("Data validated against schema: %s", schema_name)


def clear_schema_cache() -> None:
    """Clear the in-memory schema cache (useful for testing)."""
    _SCHEMA_CACHE.clear()
    logger.debug("Schema cache cleared")
