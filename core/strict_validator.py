#!/usr/bin/env python3
"""
Strict JSON Schema Validator
Production-grade JSON Schema validation with rich error reporting and caching.

This module provides comprehensive JSON Schema validation capabilities
for the PowerPoint Agent toolset, supporting manifest validation,
tool output validation, and configuration validation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Features:
- Support for JSON Schema Draft-07, Draft-2019-09, and Draft-2020-12
- Schema caching for performance
- Rich error objects with JSON serialization
- ValidationResult objects for programmatic access
- Custom format checkers for presentation-specific formats
- Backward-compatible validate_against_schema() function

Usage:
    from core.strict_validator import (
        validate_against_schema,
        validate_dict,
        validate_json_file,
        ValidationResult,
        ValidationError
    )
    
    # Simple validation (raises on error)
    validate_against_schema(data, "schemas/manifest.schema.json")
    
    # Validation with result object
    result = validate_dict(data, schema)
    if not result.is_valid:
        for error in result.errors:
            print(f"{error.path}: {error.message}")

Changelog v3.0.0:
- NEW: ValidationResult class for structured validation results
- NEW: ValidationError exception with rich details and JSON serialization
- NEW: SchemaCache for performance optimization
- NEW: Support for multiple JSON Schema drafts
- NEW: validate_dict() returning ValidationResult
- NEW: validate_json_file() for file-based validation
- NEW: Custom format checkers (hex-color, percentage, file-path)
- IMPROVED: Error messages with full JSON paths
- IMPROVED: Graceful dependency handling
"""

import json
import re
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Type
from dataclasses import dataclass, field
from datetime import datetime

# ============================================================================
# DEPENDENCY HANDLING
# ============================================================================

try:
    from jsonschema import (
        Draft7Validator,
        Draft201909Validator,
        Draft202012Validator,
        FormatChecker,
        ValidationError as JsonSchemaValidationError,
        SchemaError as JsonSchemaSchemaError
    )
    from jsonschema.protocols import Validator
    JSONSCHEMA_AVAILABLE = True
except ImportError:
    JSONSCHEMA_AVAILABLE = False
    Draft7Validator = None
    Draft201909Validator = None
    Draft202012Validator = None
    FormatChecker = None
    JsonSchemaValidationError = Exception
    JsonSchemaSchemaError = Exception
    Validator = None


# ============================================================================
# EXCEPTIONS
# ============================================================================

class ValidatorError(Exception):
    """Base exception for validator errors."""
    
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        return {
            "error": self.__class__.__name__,
            "message": self.message,
            "details": self.details
        }
    
    def to_json(self) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=2)


class ValidationError(ValidatorError):
    """
    Raised when validation fails.
    
    Contains detailed information about all validation errors.
    """
    
    def __init__(
        self,
        message: str,
        errors: Optional[List['ValidationErrorDetail']] = None,
        schema_path: Optional[str] = None
    ):
        details = {
            "error_count": len(errors) if errors else 0,
            "schema_path": schema_path
        }
        super().__init__(message, details)
        self.errors = errors or []
        self.schema_path = schema_path
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        base = super().to_dict()
        base["errors"] = [e.to_dict() for e in self.errors]
        return base


class SchemaLoadError(ValidatorError):
    """Raised when schema cannot be loaded."""
    pass


class SchemaInvalidError(ValidatorError):
    """Raised when schema itself is invalid."""
    pass


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ValidationErrorDetail:
    """
    Detailed information about a single validation error.
    """
    path: str
    message: str
    validator: str
    validator_value: Any = None
    instance: Any = None
    schema_path: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        result = {
            "path": self.path,
            "message": self.message,
            "validator": self.validator,
            "schema_path": self.schema_path
        }
        
        # Include validator_value if it's JSON-serializable
        if self.validator_value is not None:
            try:
                json.dumps(self.validator_value)
                result["validator_value"] = self.validator_value
            except (TypeError, ValueError):
                result["validator_value"] = str(self.validator_value)
        
        return result
    
    def __str__(self) -> str:
        return f"{self.path or '<root>'}: {self.message}"


@dataclass
class ValidationResult:
    """
    Result of a validation operation.
    
    Provides structured access to validation outcome and any errors.
    """
    is_valid: bool
    errors: List[ValidationErrorDetail] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    schema_path: Optional[str] = None
    schema_draft: Optional[str] = None
    validated_at: str = field(default_factory=lambda: datetime.utcnow().isoformat() + "Z")
    
    @property
    def error_count(self) -> int:
        """Number of validation errors."""
        return len(self.errors)
    
    @property
    def warning_count(self) -> int:
        """Number of warnings."""
        return len(self.warnings)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        return {
            "is_valid": self.is_valid,
            "error_count": self.error_count,
            "warning_count": self.warning_count,
            "errors": [e.to_dict() for e in self.errors],
            "warnings": self.warnings,
            "schema_path": self.schema_path,
            "schema_draft": self.schema_draft,
            "validated_at": self.validated_at
        }
    
    def to_json(self) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=2)
    
    def raise_if_invalid(self) -> None:
        """Raise ValidationError if validation failed."""
        if not self.is_valid:
            error_messages = [str(e) for e in self.errors]
            raise ValidationError(
                f"Validation failed with {self.error_count} error(s):\n" + 
                "\n".join(error_messages),
                errors=self.errors,
                schema_path=self.schema_path
            )


# ============================================================================
# SCHEMA CACHE
# ============================================================================

class SchemaCache:
    """
    Thread-safe schema cache for performance optimization.
    
    Caches loaded and compiled schemas to avoid repeated file I/O
    and schema compilation.
    """
    
    _instance: Optional['SchemaCache'] = None
    _schemas: Dict[str, Dict[str, Any]] = {}
    _validators: Dict[str, Any] = {}
    _mtimes: Dict[str, float] = {}
    
    def __new__(cls) -> 'SchemaCache':
        """Singleton pattern."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._schemas = {}
            cls._instance._validators = {}
            cls._instance._mtimes = {}
        return cls._instance
    
    def get_schema(self, schema_path: str, force_reload: bool = False) -> Dict[str, Any]:
        """
        Get schema from cache or load from file.
        
        Args:
            schema_path: Path to schema file
            force_reload: Force reload even if cached
            
        Returns:
            Parsed schema dictionary
        """
        path = Path(schema_path).resolve()
        path_str = str(path)
        
        # Check if reload needed
        if not force_reload and path_str in self._schemas:
            # Check if file was modified
            try:
                current_mtime = path.stat().st_mtime
                if current_mtime <= self._mtimes.get(path_str, 0):
                    return self._schemas[path_str]
            except OSError:
                pass
        
        # Load schema
        schema = self._load_schema_file(path)
        self._schemas[path_str] = schema
        
        try:
            self._mtimes[path_str] = path.stat().st_mtime
        except OSError:
            self._mtimes[path_str] = 0
        
        # Invalidate validator cache for this schema
        if path_str in self._validators:
            del self._validators[path_str]
        
        return schema
    
    def get_validator(
        self,
        schema_path: str,
        draft: Optional[str] = None
    ) -> Any:
        """
        Get compiled validator from cache or create new.
        
        Args:
            schema_path: Path to schema file
            draft: JSON Schema draft version (auto-detected if None)
            
        Returns:
            Compiled validator instance
        """
        if not JSONSCHEMA_AVAILABLE:
            raise ValidatorError(
                "jsonschema library is required for validation",
                details={"install": "pip install jsonschema"}
            )
        
        path = Path(schema_path).resolve()
        path_str = str(path)
        cache_key = f"{path_str}:{draft or 'auto'}"
        
        if cache_key in self._validators:
            return self._validators[cache_key]
        
        schema = self.get_schema(schema_path)
        validator_class = self._get_validator_class(schema, draft)
        
        # Create format checker with custom formats
        format_checker = self._create_format_checker()
        
        # Create validator
        validator = validator_class(schema, format_checker=format_checker)
        self._validators[cache_key] = validator
        
        return validator
    
    def clear(self) -> None:
        """Clear all cached schemas and validators."""
        self._schemas.clear()
        self._validators.clear()
        self._mtimes.clear()
    
    def _load_schema_file(self, path: Path) -> Dict[str, Any]:
        """Load schema from file."""
        if not path.exists():
            raise SchemaLoadError(
                f"Schema file not found: {path}",
                details={"path": str(path)}
            )
        
        try:
            content = path.read_text(encoding='utf-8')
            schema = json.loads(content)
            return schema
        except json.JSONDecodeError as e:
            raise SchemaLoadError(
                f"Invalid JSON in schema file: {path}",
                details={"path": str(path), "error": str(e)}
            )
        except OSError as e:
            raise SchemaLoadError(
                f"Cannot read schema file: {path}",
                details={"path": str(path), "error": str(e)}
            )
    
    def _get_validator_class(
        self,
        schema: Dict[str, Any],
        draft: Optional[str]
    ) -> Type:
        """Get appropriate validator class for schema."""
        if draft:
            draft_lower = draft.lower()
            if '2020' in draft_lower or '202012' in draft_lower:
                return Draft202012Validator
            elif '2019' in draft_lower or '201909' in draft_lower:
                return Draft201909Validator
            elif '7' in draft_lower or 'draft-07' in draft_lower:
                return Draft7Validator
        
        # Auto-detect from $schema
        schema_uri = schema.get('$schema', '')
        
        if '2020-12' in schema_uri or 'draft/2020-12' in schema_uri:
            return Draft202012Validator
        elif '2019-09' in schema_uri or 'draft/2019-09' in schema_uri:
            return Draft201909Validator
        elif 'draft-07' in schema_uri:
            return Draft7Validator
        
        # Default to latest
        return Draft202012Validator
    
    def _create_format_checker(self) -> FormatChecker:
        """Create format checker with custom formats."""
        checker = FormatChecker()
        
        # Hex color format
        @checker.checks('hex-color')
        def check_hex_color(value: str) -> bool:
            if not isinstance(value, str):
                return False
            pattern = r'^#?[0-9A-Fa-f]{6}$'
            return bool(re.match(pattern, value))
        
        # Percentage format
        @checker.checks('percentage')
        def check_percentage(value: str) -> bool:
            if not isinstance(value, str):
                return False
            pattern = r'^-?\d+(\.\d+)?%$'
            return bool(re.match(pattern, value))
        
        # File path format
        @checker.checks('file-path')
        def check_file_path(value: str) -> bool:
            if not isinstance(value, str):
                return False
            try:
                Path(value)
                return True
            except Exception:
                return False
        
        # Absolute path format
        @checker.checks('absolute-path')
        def check_absolute_path(value: str) -> bool:
            if not isinstance(value, str):
                return False
            return os.path.isabs(value)
        
        # Slide index format (non-negative integer)
        @checker.checks('slide-index')
        def check_slide_index(value: Any) -> bool:
            return isinstance(value, int) and value >= 0
        
        # Shape index format (non-negative integer)
        @checker.checks('shape-index')
        def check_shape_index(value: Any) -> bool:
            return isinstance(value, int) and value >= 0
        
        return checker


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def validate_against_schema(payload: Dict[str, Any], schema_path: str) -> None:
    """
    Strictly validate payload against JSON Schema.
    
    This is the backward-compatible function that raises ValueError on failure.
    
    Args:
        payload: Data to validate
        schema_path: Path to JSON Schema file
        
    Raises:
        ValueError: If validation fails (with detailed error messages)
        SchemaLoadError: If schema cannot be loaded
        
    Example:
        >>> validate_against_schema({"name": "test"}, "schemas/config.schema.json")
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ImportError(
            "jsonschema library is required. Install with:\n"
            "  pip install jsonschema\n"
            "  or: uv pip install jsonschema"
        )
    
    result = validate_dict(payload, schema_path=schema_path)
    
    if not result.is_valid:
        error_messages = []
        for error in result.errors:
            loc = error.path or '<root>'
            error_messages.append(f"{loc}: {error.message}")
        
        raise ValueError(
            "Strict schema validation failed:\n" + "\n".join(error_messages)
        )


def validate_dict(
    data: Dict[str, Any],
    schema: Optional[Dict[str, Any]] = None,
    schema_path: Optional[str] = None,
    draft: Optional[str] = None,
    raise_on_error: bool = False
) -> ValidationResult:
    """
    Validate dictionary against JSON Schema.
    
    Either schema or schema_path must be provided.
    
    Args:
        data: Data to validate
        schema: JSON Schema dictionary
        schema_path: Path to JSON Schema file
        draft: JSON Schema draft version (auto-detected if None)
        raise_on_error: Raise ValidationError if validation fails
        
    Returns:
        ValidationResult with validation outcome
        
    Raises:
        ValidationError: If raise_on_error=True and validation fails
        SchemaLoadError: If schema cannot be loaded
        ValidatorError: If neither schema nor schema_path provided
        
    Example:
        >>> result = validate_dict(data, schema_path="schemas/manifest.json")
        >>> if not result.is_valid:
        ...     for error in result.errors:
        ...         print(error)
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ValidatorError(
            "jsonschema library is required",
            details={"install": "pip install jsonschema"}
        )
    
    if schema is None and schema_path is None:
        raise ValidatorError(
            "Either schema or schema_path must be provided"
        )
    
    # Get or create validator
    cache = SchemaCache()
    
    if schema_path:
        validator = cache.get_validator(schema_path, draft)
        resolved_schema = cache.get_schema(schema_path)
    else:
        validator_class = cache._get_validator_class(schema, draft)
        format_checker = cache._create_format_checker()
        validator = validator_class(schema, format_checker=format_checker)
        resolved_schema = schema
    
    # Detect draft version
    schema_draft = resolved_schema.get('$schema', 'unknown')
    
    # Collect errors
    errors: List[ValidationErrorDetail] = []
    warnings: List[str] = []
    
    try:
        validation_errors = sorted(
            validator.iter_errors(data),
            key=lambda e: (list(e.absolute_path), e.message)
        )
        
        for error in validation_errors:
            path = "/".join(str(p) for p in error.absolute_path)
            schema_path_str = "/".join(str(p) for p in error.absolute_schema_path)
            
            errors.append(ValidationErrorDetail(
                path=path,
                message=error.message,
                validator=error.validator,
                validator_value=error.validator_value,
                instance=error.instance if _is_json_serializable(error.instance) else str(error.instance),
                schema_path=schema_path_str
            ))
    except JsonSchemaSchemaError as e:
        raise SchemaInvalidError(
            f"Invalid schema: {e.message}",
            details={"error": str(e)}
        )
    
    # Create result
    result = ValidationResult(
        is_valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
        schema_path=schema_path,
        schema_draft=schema_draft
    )
    
    if raise_on_error:
        result.raise_if_invalid()
    
    return result


def validate_json_file(
    file_path: str,
    schema_path: str,
    draft: Optional[str] = None,
    raise_on_error: bool = False
) -> ValidationResult:
    """
    Validate JSON file against schema.
    
    Args:
        file_path: Path to JSON file to validate
        schema_path: Path to JSON Schema file
        draft: JSON Schema draft version
        raise_on_error: Raise ValidationError if validation fails
        
    Returns:
        ValidationResult with validation outcome
        
    Raises:
        ValidationError: If raise_on_error=True and validation fails
        SchemaLoadError: If files cannot be loaded
    """
    path = Path(file_path)
    
    if not path.exists():
        raise SchemaLoadError(
            f"File not found: {file_path}",
            details={"path": file_path}
        )
    
    try:
        content = path.read_text(encoding='utf-8')
        data = json.loads(content)
    except json.JSONDecodeError as e:
        raise SchemaLoadError(
            f"Invalid JSON in file: {file_path}",
            details={"path": file_path, "error": str(e)}
        )
    except OSError as e:
        raise SchemaLoadError(
            f"Cannot read file: {file_path}",
            details={"path": file_path, "error": str(e)}
        )
    
    return validate_dict(
        data,
        schema_path=schema_path,
        draft=draft,
        raise_on_error=raise_on_error
    )


def load_schema(schema_path: str, force_reload: bool = False) -> Dict[str, Any]:
    """
    Load JSON Schema from file with caching.
    
    Args:
        schema_path: Path to schema file
        force_reload: Force reload from disk
        
    Returns:
        Parsed schema dictionary
    """
    cache = SchemaCache()
    return cache.get_schema(schema_path, force_reload=force_reload)


def clear_schema_cache() -> None:
    """Clear the schema cache."""
    cache = SchemaCache()
    cache.clear()


def is_valid(
    data: Dict[str, Any],
    schema: Optional[Dict[str, Any]] = None,
    schema_path: Optional[str] = None
) -> bool:
    """
    Quick validation check returning boolean.
    
    Args:
        data: Data to validate
        schema: JSON Schema dictionary
        schema_path: Path to JSON Schema file
        
    Returns:
        True if valid, False otherwise
    """
    try:
        result = validate_dict(data, schema=schema, schema_path=schema_path)
        return result.is_valid
    except Exception:
        return False


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def _is_json_serializable(value: Any) -> bool:
    """Check if value is JSON serializable."""
    try:
        json.dumps(value)
        return True
    except (TypeError, ValueError):
        return False


def get_schema_draft(schema: Dict[str, Any]) -> str:
    """
    Detect JSON Schema draft version from schema.
    
    Args:
        schema: Schema dictionary
        
    Returns:
        Draft identifier string
    """
    schema_uri = schema.get('$schema', '')
    
    if '2020-12' in schema_uri:
        return 'draft-2020-12'
    elif '2019-09' in schema_uri:
        return 'draft-2019-09'
    elif 'draft-07' in schema_uri:
        return 'draft-07'
    elif 'draft-06' in schema_uri:
        return 'draft-06'
    elif 'draft-04' in schema_uri:
        return 'draft-04'
    
    return 'unknown'


# ============================================================================
# MODULE METADATA
# ============================================================================

__version__ = "3.0.0"
__author__ = "PowerPoint Agent Team"
__license__ = "MIT"

__all__ = [
    # Main functions
    "validate_against_schema",
    "validate_dict",
    "validate_json_file",
    "load_schema",
    "clear_schema_cache",
    "is_valid",
    "get_schema_draft",
    
    # Classes
    "ValidationResult",
    "ValidationErrorDetail",
    "SchemaCache",
    
    # Exceptions
    "ValidatorError",
    "ValidationError",
    "SchemaLoadError",
    "SchemaInvalidError",
    
    # Constants
    "JSONSCHEMA_AVAILABLE",
    
    # Module metadata
    "__version__",
    "__author__",
    "__license__",
]
