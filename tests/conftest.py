"""
Pytest configuration and fixtures for PPT MCP Server tests.
"""

import pytest


@pytest.fixture(autouse=True)
def reset_ppt_automation():
    """Reset the ppt_automation presentations before each test."""
    from main import ppt_automation

    # Store original presentations
    original_presentations = ppt_automation.presentations.copy()

    yield

    # Restore original state (clean up test presentations)
    ppt_automation.presentations.clear()
    ppt_automation.presentations.update(original_presentations)
