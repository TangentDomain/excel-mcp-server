#!/usr/bin/env python3
"""
Setup pre-commit hooks for Excel MCP Server project
"""

import subprocess
import sys
from pathlib import Path


def check_precommit_installed():
    """Check if pre-commit is installed"""

    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "show", "pre-commit"],
            capture_output=True,
            text=True
        )
        return result.returncode == 0
    except Exception:
        return False


def safe_print(message, fallback=None):
    """Print with fallback for encoding issues"""
    try:
        print(message)
    except UnicodeEncodeError:
        if fallback:
            print(fallback)
        else:
            # Remove emoji characters as fallback
            import re
            clean_message = re.sub(r'[^\x00-\x7F]+', '', message)
            print(clean_message)


def install_precommit():
    """Install pre-commit if not already installed"""

    if not check_precommit_installed():
        safe_print("📦 Installing pre-commit...", "Installing pre-commit...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "pre-commit"],
                check=True
            )
            safe_print("✅ pre-commit installed successfully", "pre-commit installed successfully")
        except subprocess.CalledProcessError as e:
            safe_print(f"❌ Failed to install pre-commit: {e}", f"Failed to install pre-commit: {e}")
            return False
    else:
        safe_print("✅ pre-commit is already installed", "pre-commit is already installed")

    return True


def setup_hooks():
    """Setup pre-commit hooks"""

    safe_print("🔧 Setting up pre-commit hooks...", "Setting up pre-commit hooks...")

    try:
        # Install the hooks
        result = subprocess.run(
            ["pre-commit", "install"],
            capture_output=True,
            text=True,
            cwd=Path.cwd()
        )

        if result.returncode == 0:
            safe_print("✅ Pre-commit hooks installed successfully", "Pre-commit hooks installed successfully")
            return True
        else:
            safe_print("❌ Failed to install pre-commit hooks:", "Failed to install pre-commit hooks:")
            print(result.stderr)
            return False

    except FileNotFoundError:
        safe_print("❌ pre-commit command not found. Please ensure it's installed and in PATH",
                  "pre-commit command not found. Please ensure it's installed and in PATH")
        return False
    except Exception as e:
        safe_print(f"❌ Error setting up pre-commit hooks: {e}", f"Error setting up pre-commit hooks: {e}")
        return False


def run_autoupdate():
    """Run pre-commit autoupdate to get latest hook versions"""

    safe_print("🔄 Running pre-commit autoupdate...", "Running pre-commit autoupdate...")
    try:
        result = subprocess.run(
            ["pre-commit", "autoupdate"],
            capture_output=True,
            text=True,
            cwd=Path.cwd()
        )

        if result.returncode == 0:
            safe_print("✅ Pre-commit hooks updated to latest versions", "Pre-commit hooks updated to latest versions")
            print(result.stdout)
        else:
            safe_print("⚠️  Some hooks could not be updated:", "Some hooks could not be updated:")
            print(result.stderr)

    except Exception as e:
        safe_print(f"⚠️  Error running autoupdate: {e}", f"Error running autoupdate: {e}")


def run_test_commit():
    """Run pre-commit on all files to test setup"""

    safe_print("🧪 Testing pre-commit setup on all files...", "Testing pre-commit setup on all files...")
    try:
        result = subprocess.run(
            ["pre-commit", "run", "--all-files"],
            capture_output=True,
            text=True,
            cwd=Path.cwd()
        )

        print(result.stdout)
        if result.stderr:
            print("Warnings/Errors:")
            print(result.stderr)

        if result.returncode == 0:
            safe_print("✅ All pre-commit hooks passed", "All pre-commit hooks passed")
            return True
        else:
            safe_print("⚠️  Some pre-commit hooks failed. Fix the issues and commit again.",
                      "Some pre-commit hooks failed. Fix the issues and commit again.")
            return False

    except Exception as e:
        safe_print(f"❌ Error running pre-commit test: {e}", f"Error running pre-commit test: {e}")
        return False


def print_usage_info():
    """Print usage information"""

    try:
        print("\n📖 Pre-commit Usage Information:")
    except UnicodeEncodeError:
        print("\nPre-commit Usage Information:")

    print("=" * 50)
    print("• Run on all files:     pre-commit run --all-files")
    print("• Run on staged files:  pre-commit run")
    print("• Update hooks:         pre-commit autoupdate")
    print("• Clean caches:         pre-commit clean")
    print("• Skip a hook:          SKIP=hook_name git commit")

    try:
        print("\n🔧 Available hooks:")
    except UnicodeEncodeError:
        print("\nAvailable hooks:")

    print("• black - Python code formatting")
    print("• isort - Import sorting")
    print("• flake8 - Code quality checks")
    print("• mypy - Type checking")
    print("• trailing-whitespace - Remove trailing whitespace")
    print("• end-of-file-fixer - Ensure files end with newline")
    print("• check-yaml - YAML syntax validation")
    print("• check-json - JSON syntax validation")
    print("• pytest-check - Run tests before commit")
    print("• check-project-structure - Validate project structure")
    print("• validate-mcp-tools - Validate MCP tool definitions")


def main():
    """Main setup function"""

    # Handle Windows console encoding issues
    try:
        print("🚀 Setting up pre-commit hooks for Excel MCP Server")
        print("=" * 55)
    except UnicodeEncodeError:
        print("Setting up pre-commit hooks for Excel MCP Server")
        print("=" * 55)

    # Step 1: Install pre-commit if needed
    if not install_precommit():
        sys.exit(1)

    # Step 2: Setup the hooks
    if not setup_hooks():
        sys.exit(1)

    # Step 3: Update to latest versions
    run_autoupdate()

    # Step 4: Test the setup (optional)
    try:
        test_input = input("\n🧪 Run pre-commit on all files to test setup? (y/N): ")
    except UnicodeEncodeError:
        test_input = input("\nRun pre-commit on all files to test setup? (y/N): ")

    if test_input.lower().startswith('y'):
        run_test_commit()

    # Step 5: Print usage information
    print_usage_info()

    safe_print("\n🎉 Pre-commit setup complete!", "\nPre-commit setup complete!")
    print("   Hooks will now run automatically before each commit.")


if __name__ == "__main__":
    main()