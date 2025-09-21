# Publishing Guide

## Publishing to PyPI

### Automatic Publishing (Recommended)

1. **Set up PyPI trusted publishing**:
   - Go to https://pypi.org/manage/account/publishing/
   - Click "Add a new pending publisher"
   - Fill in the form exactly as follows:
     - **PyPI Project Name**: `mcp-powerpoint`
     - **Owner**: `CannonJunior` (your GitHub username)
     - **Repository name**: `mcp-powerpoint`
     - **Workflow name**: `publish.yml`
     - **Environment name**: (leave blank)
   - Click "Add"

2. **Create a release**:

   **Option A: Via GitHub Web Interface (Recommended)**
   - Go to https://github.com/CannonJunior/mcp-powerpoint/releases
   - Look for and click one of these buttons:
     - "Create a new release" (if releases exist)
     - "Create release" (if no releases exist yet)
     - "Draft a new release" (alternative text)
   - If you don't see any of these, go to the main repo page and look for a "Releases" section on the right side, then click "Create a new release"
   - Click "Choose a tag" and type `v1.0.0` (this will create the tag)
   - Set "Release title" to `v1.0.0`
   - In the description, add release notes like:
     ```
     ## MCP PowerPoint Tools v1.0.0

     First stable release with uvx installation support.

     ### Features
     - PowerPoint to JSON conversion
     - JSON to PowerPoint reconstruction
     - LLM-powered shape naming via Ollama
     - uvx installation support
     ```
   - Click "Publish release"

   **Option B: Via Command Line**
   ```bash
   # Create and push the tag
   git tag v1.0.0
   git push origin v1.0.0

   # Then go to GitHub and create the release from the tag
   # Or use GitHub CLI:
   gh release create v1.0.0 --title "v1.0.0" --notes "First stable release"
   ```

3. **GitHub Actions will automatically**:
   - Build the package
   - Publish to PyPI
   - Users can then run: `uvx mcp-powerpoint --server powerpoint`

### Manual Publishing

1. **Install publishing tools**:
   ```bash
   uv add --dev twine
   ```

2. **Build and publish**:
   ```bash
   uv build
   uv run twine upload dist/*
   ```

## Current Working Installation Methods

Until published to PyPI, users can install using these methods:

### Method 1: From GitHub (Recommended)
```json
{
  "mcpServers": {
    "mcp-powerpoint": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/your-username/mcp-powerpoint.git", "mcp-powerpoint", "--server", "powerpoint"]
    }
  }
}
```

### Method 2: From Local Directory
```json
{
  "mcpServers": {
    "mcp-powerpoint": {
      "command": "uvx",
      "args": ["--from", "/path/to/mcp-powerpoint", "mcp-powerpoint", "--server", "powerpoint"]
    }
  }
}
```

### Method 3: From Wheel File
```json
{
  "mcpServers": {
    "mcp-powerpoint": {
      "command": "uvx",
      "args": ["--from", "/path/to/mcp_powerpoint-1.0.0-py3-none-any.whl", "mcp-powerpoint", "--server", "powerpoint"]
    }
  }
}
```