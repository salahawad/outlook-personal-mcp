import logging

from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.server import configure_logging


def test_quiets_noisy_loggers_by_default():
    # Simulate the noisy default (FastMCP/root configures INFO).
    for name in ("httpx", "httpcore", "msal", "urllib3"):
        logging.getLogger(name).setLevel(logging.INFO)
    configure_logging(Settings(client_id="abc"))
    for name in ("httpx", "httpcore", "msal", "urllib3"):
        assert logging.getLogger(name).level == logging.WARNING


def test_debug_enables_info_logging():
    configure_logging(Settings(client_id="abc", debug=True))
    assert logging.getLogger("httpx").level == logging.INFO
    assert logging.getLogger("msal").level == logging.INFO
