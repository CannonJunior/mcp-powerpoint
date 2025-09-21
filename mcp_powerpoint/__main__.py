#!/usr/bin/env python3
"""
Main entry point for mcp-powerpoint package.
Allows launching different MCP servers via command line arguments.
"""

import sys
import argparse
from .powerpoint_server import main as powerpoint_main
from .shape_naming_server import main as shape_naming_main


def main():
    parser = argparse.ArgumentParser(description="MCP PowerPoint Tools")
    parser.add_argument(
        "--server",
        choices=["powerpoint", "shape-naming"],
        required=True,
        help="Which server to launch"
    )

    args = parser.parse_args()

    if args.server == "powerpoint":
        powerpoint_main()
    elif args.server == "shape-naming":
        shape_naming_main()
    else:
        print(f"Unknown server: {args.server}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()