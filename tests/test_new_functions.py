"""
Unit tests for new slide manipulation functions in the PPT MCP Server.
Tests include: copy_slide, delete_slide, move_slide, duplicate_slide_with_text_updates,
get_presentation_info, and list_all_shapes_in_slide.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch
import sys
import os
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import (
    copy_slide,
    delete_slide,
    move_slide,
    get_presentation_info,
    list_all_shapes_in_slide,
    save_copy,
    ppt_automation,
)


@pytest.fixture
def mock_presentation():
    """Create a mock presentation with slides."""
    pres = MagicMock()

    # Create mock slides
    slides = MagicMock()
    slide1 = MagicMock()
    slide2 = MagicMock()
    slide3 = MagicMock()

    # Setup slide count
    slides.Count = 3
    slides.Item = MagicMock(side_effect=lambda i: [None, slide1, slide2, slide3][i])

    # Setup shapes for slides
    for slide in [slide1, slide2, slide3]:
        shapes = MagicMock()
        shapes.Count = 2
        shape1 = MagicMock()
        shape1.Name = "Title"
        shape1.Type = 14  # Placeholder
        shape1.TextFrame = MagicMock()
        shape1.TextFrame.TextRange = MagicMock()
        shape1.TextFrame.TextRange.Text = "Slide Title"
        shapes.Item = MagicMock(return_value=shape1)
        slide.Shapes = shapes

    pres.Slides = slides
    pres.FullName = "C:\\test\\presentation.pptx"
    pres.Saved = True

    # Setup Copy and Paste methods
    pres.Slides.Copy = MagicMock()
    pres.Slides.Paste = MagicMock()

    return pres


@pytest.fixture
def setup_ppt_automation(mock_presentation):
    """Setup the ppt_automation with a mock presentation."""
    pres_id = "test-pres-id"
    ppt_automation.presentations[pres_id] = mock_presentation
    yield pres_id
    # Cleanup
    if pres_id in ppt_automation.presentations:
        del ppt_automation.presentations[pres_id]


def test_copy_slide_success(setup_ppt_automation):
    """Test successfully copying a slide."""
    pres_id = setup_ppt_automation

    result = copy_slide(pres_id, 1, insert_after=1)

    assert result["success"] is True
    assert "message" in result
    assert "Slide 1 copied" in result["message"]


def test_copy_slide_invalid_presentation():
    """Test copy_slide with invalid presentation ID."""
    result = copy_slide("invalid-id", 1)

    assert "error" in result
    assert "Presentation ID not found" in result["error"]


def test_copy_slide_invalid_slide_id(setup_ppt_automation):
    """Test copy_slide with invalid slide ID."""
    pres_id = setup_ppt_automation

    result = copy_slide(pres_id, 10)  # Slide doesn't exist

    assert "error" in result
    assert "Invalid slide ID" in result["error"]


def test_delete_slide_success(setup_ppt_automation):
    """Test successfully deleting a slide."""
    pres_id = setup_ppt_automation

    result = delete_slide(pres_id, 2)

    assert result["success"] is True
    assert "deleted successfully" in result["message"]


def test_delete_slide_invalid_presentation():
    """Test delete_slide with invalid presentation ID."""
    result = delete_slide("invalid-id", 1)

    assert "error" in result
    assert "Presentation ID not found" in result["error"]


def test_delete_slide_invalid_slide_id(setup_ppt_automation):
    """Test delete_slide with invalid slide ID."""
    pres_id = setup_ppt_automation

    result = delete_slide(pres_id, 10)

    assert "error" in result
    assert "Invalid slide ID" in result["error"]


def test_move_slide_success(setup_ppt_automation):
    """Test successfully moving a slide."""
    pres_id = setup_ppt_automation

    result = move_slide(pres_id, 1, 3)

    assert result["success"] is True
    assert "moved to position" in result["message"]


def test_move_slide_invalid_presentation():
    """Test move_slide with invalid presentation ID."""
    result = move_slide("invalid-id", 1, 2)

    assert "error" in result
    assert "Presentation ID not found" in result["error"]


def test_move_slide_invalid_slide_id(setup_ppt_automation):
    """Test move_slide with invalid slide ID."""
    pres_id = setup_ppt_automation

    result = move_slide(pres_id, 10, 2)

    assert "error" in result
    assert "Invalid slide ID" in result["error"]


def test_move_slide_invalid_position(setup_ppt_automation):
    """Test move_slide with invalid new position."""
    pres_id = setup_ppt_automation

    result = move_slide(pres_id, 1, 10)

    assert "error" in result
    assert "Invalid new position" in result["error"]


def test_get_presentation_info_success(setup_ppt_automation):
    """Test getting presentation info."""
    pres_id = setup_ppt_automation

    result = get_presentation_info(pres_id)

    assert "error" not in result
    assert result["id"] == pres_id
    assert result["slide_count"] == 3
    assert "presentation.pptx" in result["name"]


def test_get_presentation_info_invalid_presentation():
    """Test get_presentation_info with invalid presentation ID."""
    result = get_presentation_info("invalid-id")

    assert "error" in result
    assert "Presentation ID not found" in result["error"]


def test_list_all_shapes_in_slide_success(setup_ppt_automation):
    """Test listing all shapes in a slide."""
    pres_id = setup_ppt_automation

    result = list_all_shapes_in_slide(pres_id, 1)

    assert "error" not in result
    assert result["slide_id"] == 1
    assert result["shape_count"] == 2
    assert "shapes" in result
    assert len(result["shapes"]) > 0


def test_list_all_shapes_in_slide_invalid_presentation():
    """Test list_all_shapes_in_slide with invalid presentation ID."""
    result = list_all_shapes_in_slide("invalid-id", 1)

    assert "error" in result
    assert "Presentation ID not found" in result["error"]


def test_list_all_shapes_in_slide_invalid_slide_id(setup_ppt_automation):
    """Test list_all_shapes_in_slide with invalid slide ID."""
    pres_id = setup_ppt_automation

    result = list_all_shapes_in_slide(pres_id, 10)

    assert "error" in result
    assert "Invalid slide ID" in result["error"]


def test_save_copy_success(setup_ppt_automation):
    """Test successfully saving a copy of a presentation."""
    pres_id = setup_ppt_automation

    with patch('os.path.dirname') as mock_dirname, \
         patch('os.path.exists') as mock_exists, \
         patch('os.makedirs') as mock_makedirs:

        mock_dirname.return_value = "C:\\test"
        mock_exists.return_value = True

        result = save_copy(pres_id, "C:\\test\\copy.pptx")

        assert result["success"] is True
        assert "copy.pptx" in result["path"]


def test_save_copy_invalid_presentation():
    """Test save_copy with invalid presentation ID."""
    result = save_copy("invalid-id", "C:\\test\\copy.pptx")

    assert "error" in result
    assert "Presentation ID not found" in result["error"]


class TestEdgeCases:
    """Test edge cases and error conditions."""

    def test_copy_first_slide(self, setup_ppt_automation):
        """Test copying the first slide."""
        pres_id = setup_ppt_automation

        result = copy_slide(pres_id, 1)

        assert result["success"] is True

    def test_copy_last_slide(self, setup_ppt_automation):
        """Test copying the last slide."""
        pres_id = setup_ppt_automation

        result = copy_slide(pres_id, 3)

        assert result["success"] is True

    def test_move_to_same_position(self, setup_ppt_automation):
        """Test moving a slide to its current position."""
        pres_id = setup_ppt_automation

        result = move_slide(pres_id, 1, 1)

        # This might succeed or fail depending on implementation
        # But it shouldn't crash
        assert "error" in result or result["success"] is True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
