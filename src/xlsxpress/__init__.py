import importlib.metadata

if __package__ is not None:
    _version = importlib.metadata.version(__package__)
else:
    _version = "unknown"

__VERSION__ = _version
