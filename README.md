# PowerPoint Automation MCP Server for Claude Desktop

This project provides a PowerPoint automation server that works with Claude Desktop via the Model Control Protocol (MCP). It allows Claude to interact with Microsoft PowerPoint, enabling tasks like creating presentations, adding slides, modifying content, and more.

## Features

- Create, open, save, and close PowerPoint presentations
- List all open presentations
- Get slide information and content
- Add new slides with different layouts
- Add text boxes to slides
- Update text content in shapes
- Set slide titles
- And more!

## Quick Start
1. Use `uvx` to run:
```bash
uvx --from https://github.com/socamalo/PPT_MCP_Server.git mcp-ppt
```

## Installation

1. Clone this repository:

2. Install dependencies:

   ```bash
   uv add fastmcp pywin32
   ```

3. Configure Claude Desktop:
   - Open Claude Desktop
   - Navigate to settings
   - Configure the MCP server as explained below

## Configuration

To configure Claude Desktop to use this MCP server, add the following to your Claude Desktop configuration file, located at `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "ppts": {
      "command": "uv",
      "args": ["run", "path/to/main.py"]
    }
  }
}
```

If you're using a virtual environment or alternative Python executable (like `uv`):

```json
{
  "mcpServers": {
    "ppts": {
      "command": "C:\\Path\\To\\Python\\Scripts\\uv.exe",
      "args": ["run", "C:\\Path\\To\\Project\\main.py"]
    }
  }
}
```

## Usage

Once configured, you can use Claude Desktop to control PowerPoint. Example interactions:

1. Initialize PowerPoint:

   ```
   Could you open PowerPoint for me?
   ```

2. Create a new presentation:

   ```
   Please create a new PowerPoint presentation.
   ```

3. Add a slide:

   ```
   Add a new slide to the presentation.
   ```

4. Add content:

   ```
   Add a text box to slide 1 with the text "Hello World".
   ```

5. Save the presentation:
   ```
   Save the presentation to C:\Users\username\Documents\presentation.pptx
   ```

## Slide Duplication Workflow

A common use case is to duplicate slides while preserving the design and updating only the content. Here's the workflow:

1. **Copy a slide with formatting preserved:**
   ```
   Copy slide 1 and insert it after slide 3.
   ```

2. **List the shapes on the new slide to identify which ones to update:**
   ```
   List all shapes in slide 4 so I can see the shape IDs and their current text.
   ```

3. **Update text in specific shapes:**
   ```
   Update the text in shape 1 of slide 4 to "New Title".
   Update the text in shape 3 of slide 4 to "New Content".
   ```

This approach is particularly useful for creating presentation templates where you want to:
- Keep the design, colors, and layout consistent
- Only change the text content
- Maintain all formatting automatically
- Have fine-grained control over which shapes to update

## Available Functions

The server provides the following PowerPoint automation functions:

### Presentation Management
- `initialize_powerpoint()`: Connect to PowerPoint and make it visible
- `get_presentations()`: List all open presentations
- `open_presentation(path)`: Open a presentation from a file
- `create_presentation()`: Create a new presentation
- `save_presentation(presentation_id, path)`: Save a presentation to disk
- `save_copy(presentation_id, path)`: Create a copy of a presentation at a new location
- `close_presentation(presentation_id, save)`: Close a presentation
- `get_presentation_info(presentation_id)`: Get metadata about a presentation

### Slide Operations
- `get_slides(presentation_id)`: Get all slides in a presentation
- `add_slide(presentation_id, layout_type)`: Add a new slide
- `copy_slide(presentation_id, slide_id, insert_after)`: Copy a slide preserving all formatting and design
- `delete_slide(presentation_id, slide_id)`: Delete a slide from a presentation
- `move_slide(presentation_id, slide_id, new_position)`: Move a slide to a new position

### Content Management
- `get_slide_text(presentation_id, slide_id)`: Get text content of a slide
- `update_text(presentation_id, slide_id, shape_id, text)`: Update text in a shape
- `add_text_box(presentation_id, slide_id, text, left, top, width, height)`: Add a text box
- `set_slide_title(presentation_id, slide_id, title)`: Set the title of a slide
- `list_all_shapes_in_slide(presentation_id, slide_id)`: List all shapes in a slide with detailed information

## Requirements

- Windows with Microsoft PowerPoint installed
- Python 3.7+
- Claude Desktop client
- `pywin32` and `fastmcp` Python packages

## Limitations

- Works only on Windows with PowerPoint installed
- The PowerPoint application will open and be visible during operations
- Limited to the capabilities exposed by the PowerPoint COM API

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

[MIT License](LICENSE)
