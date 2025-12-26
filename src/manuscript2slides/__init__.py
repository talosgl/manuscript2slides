try:
    from importlib.metadata import version

    __version__ = version("manuscript2slides")
except Exception:
    __version__ = "unknown"
