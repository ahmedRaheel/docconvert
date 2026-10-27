class DocConvertError(Exception):
    """Base class for all docconvert errors."""


class UnsupportedConversionError(DocConvertError):
    """The requested source/target format combination is not supported."""


class ConversionError(DocConvertError):
    """A conversion started but could not be completed."""


class MissingDependencyError(DocConvertError):
    """An optional dependency or external program is not available."""
