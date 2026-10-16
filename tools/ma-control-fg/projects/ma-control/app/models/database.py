# -*- coding: utf-8 -*-
"""Database initialization and SQLite manager for MA CONTROL."""

# -*- coding: utf-8 -*-
"""Re-export database module from app.core.database for backward compatibility."""

from app.core.database import get_connection, init_db

__all__ = ["get_connection", "init_db"]
