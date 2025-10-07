
import logging
from typing import Optional

def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """Configure basic structured logging for the app.

    Args:
        level: Logging level (e.g., logging.INFO, logging.DEBUG).
    Returns:
        A module-level logger named 'clean_switch'.
    """
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    return logging.getLogger("clean_switch")
