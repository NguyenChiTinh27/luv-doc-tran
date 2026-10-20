"""
src/engine/phase3_bilingual/cost_guard.py
──────────────────────────────────────────
Zero-cost runtime interceptor and context manager for Phase 3 (R9).

Strictly blocks all outgoing HTTP/HTTPS network calls to paid translation
endpoints (e.g. api.deepseek.com, api.openai.com) during testing, development,
and offline conversion pipelines to protect the user's API budget.
"""

import os
import sys
import functools
from typing import List, Optional, Callable, Any
from contextlib import contextmanager


class ApiCostViolationError(Exception):
    """
    Raised immediately when an unauthorized or accidental external paid API call
    is intercepted by ZeroCostGuard.
    """
    pass


class ZeroCostGuard:
    """
    Runtime security and cost guard that intercepts outgoing HTTP/HTTPS calls
    via requests, httpx, and urllib.request.

    Can be used:
      1. As a context manager:
         with ZeroCostGuard():
             # calls to api.deepseek.com or api.openai.com will raise ApiCostViolationError
             ...

      2. Globally enabled:
         ZeroCostGuard.enable()
         ...
         ZeroCostGuard.disable()

      3. As a function decorator:
         @ZeroCostGuard.protect
         def test_something():
             ...
    """

    DEFAULT_BLOCKED_HOSTS: List[str] = [
        "api.deepseek.com",
        "api.openai.com",
        "openai.com",
        "api.anthropic.com",
    ]

    _is_active: bool = False
    _original_requests_send = None
    _original_httpx_send = None
    _original_urllib_urlopen = None
    _blocked_hosts: List[str] = list(DEFAULT_BLOCKED_HOSTS)

    @classmethod
    def is_active(cls) -> bool:
        """Returns True if the guard interceptors are currently active or running in offline/zero-cost mode."""
        return cls._is_active or os.environ.get("ZERO_COST_GUARD") == "1" or os.environ.get("OFFLINE_TEST_MODE") == "1"

    @classmethod
    def set_blocked_hosts(cls, hosts: List[str]) -> None:
        """Configures or extends the list of blocked domain substrings."""
        cls._blocked_hosts = list(hosts)

    @classmethod
    def reset_blocked_hosts(cls) -> None:
        """Resets blocked hosts to default list."""
        cls._blocked_hosts = list(cls.DEFAULT_BLOCKED_HOSTS)

    @classmethod
    def _check_url(cls, url: str) -> None:
        """Inspects target URL against blocked host list."""
        if not url:
            return
        url_lower = str(url).lower()
        for host in cls._blocked_hosts:
            if host in url_lower:
                raise ApiCostViolationError(
                    f"Prohibited paid API call to {url} blocked by ZeroCostGuard! (R9 Cost Constraint)"
                )

    @classmethod
    def enable(cls, blocked_hosts: Optional[List[str]] = None) -> None:
        """
        Installs monkey-patches into requests, httpx, and urllib.request.
        """
        if blocked_hosts is not None:
            cls.set_blocked_hosts(blocked_hosts)

        if cls._is_active:
            return
        cls._is_active = True

        # 1. Patch requests.Session.send
        try:
            import requests
            cls._original_requests_send = requests.Session.send

            def guarded_requests_send(session_self, request, **kwargs):
                url = str(getattr(request, "url", ""))
                cls._check_url(url)
                return cls._original_requests_send(session_self, request, **kwargs)

            requests.Session.send = guarded_requests_send
        except ImportError:
            pass

        # 2. Patch httpx.Client.send
        try:
            import httpx
            cls._original_httpx_send = httpx.Client.send

            def guarded_httpx_send(client_self, request, **kwargs):
                url = str(getattr(request, "url", ""))
                cls._check_url(url)
                return cls._original_httpx_send(client_self, request, **kwargs)

            httpx.Client.send = guarded_httpx_send
        except ImportError:
            pass

        # 3. Patch urllib.request.urlopen
        try:
            import urllib.request
            cls._original_urllib_urlopen = urllib.request.urlopen

            def guarded_urlopen(url, *args, **kwargs):
                url_str = str(getattr(url, "full_url", url))
                cls._check_url(url_str)
                return cls._original_urllib_urlopen(url, *args, **kwargs)

            urllib.request.urlopen = guarded_urlopen
        except Exception:
            pass

    @classmethod
    def disable(cls) -> None:
        """
        Restores original networking functions and disables interception.
        """
        if not cls._is_active:
            return
        cls._is_active = False

        try:
            import requests
            if cls._original_requests_send is not None:
                requests.Session.send = cls._original_requests_send
                cls._original_requests_send = None
        except ImportError:
            pass

        try:
            import httpx
            if cls._original_httpx_send is not None:
                httpx.Client.send = cls._original_httpx_send
                cls._original_httpx_send = None
        except ImportError:
            pass

        try:
            import urllib.request
            if cls._original_urllib_urlopen is not None:
                urllib.request.urlopen = cls._original_urllib_urlopen
                cls._original_urllib_urlopen = None
        except Exception:
            pass

    def __enter__(self):
        self.enable()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disable()

    @classmethod
    def protect(cls, func: Callable) -> Callable:
        """Decorator to wrap a test or function inside ZeroCostGuard context."""
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            with cls():
                return func(*args, **kwargs)
        return wrapper
