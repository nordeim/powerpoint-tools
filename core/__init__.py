"""
PowerPoint Agent Core Package
"""

from .powerpoint_agent_core import (
    # Core class
    PowerPointAgent,
    
    # Exceptions
    PowerPointAgentError,
    SlideNotFoundError,
    LayoutNotFoundError,
    ImageNotFoundError,
    InvalidPositionError,
    TemplateError,
    ThemeError,
    AccessibilityError,
    AssetValidationError,
    FileLockError,
    
    # Helpers
    Position,
    Size,
    ColorHelper,
    TemplateProfile,
    AccessibilityChecker,
    AssetValidator,
    
    # Enums
    ShapeType,
    ChartType,
    TextAlignment,
    VerticalAlignment,
    BulletStyle,
    ImageFormat,
    ExportFormat,
    
    # Constants
    SLIDE_WIDTH_INCHES,
    SLIDE_HEIGHT_INCHES,
    ANCHOR_POINTS,
    CORPORATE_COLORS,
    STANDARD_FONTS,
)

__version__ = "1.0.0"
__all__ = [
    "PowerPointAgent",
    "PowerPointAgentError",
    "SlideNotFoundError",
    "LayoutNotFoundError",
    "ImageNotFoundError",
    "InvalidPositionError",
    "TemplateError",
    "ThemeError",
    "AccessibilityError",
    "AssetValidationError",
    "FileLockError",
    "Position",
    "Size",
    "ColorHelper",
    "TemplateProfile",
    "AccessibilityChecker",
    "AssetValidator",
    "ShapeType",
    "ChartType",
    "TextAlignment",
    "VerticalAlignment",
    "BulletStyle",
    "ImageFormat",
    "ExportFormat",
    "SLIDE_WIDTH_INCHES",
    "SLIDE_HEIGHT_INCHES",
    "ANCHOR_POINTS",
    "CORPORATE_COLORS",
    "STANDARD_FONTS",
]
