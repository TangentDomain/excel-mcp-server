#!/usr/bin/env python3
"""
Check test coverage meets minimum threshold
"""

import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


def run_coverage_check():
    """Run coverage check and parse results"""

    print("🔍 Running test coverage check...")

    try:
        # Run coverage report in XML format
        result = subprocess.run(
            [
                sys.executable, "-m", "pytest",
                "--cov=src",
                "--cov-report=xml",
                "--cov-report=term-missing",
                "tests/",
                "-q"
            ],
            capture_output=True,
            text=True,
            cwd=Path.cwd()
        )

        if result.returncode != 0:
            print("❌ Tests failed, cannot check coverage")
            print(f"Error: {result.stderr}")
            return False

        # Parse coverage XML report
        coverage_file = Path("coverage.xml")
        if not coverage_file.exists():
            print("❌ coverage.xml not found")
            return False

        tree = ET.parse(coverage_file)
        root = tree.getroot()

        # Find overall coverage
        coverage = root.find("coverage")
        if coverage is None:
            print("❌ Could not find coverage data in XML")
            return False

        line_rate = float(coverage.get("line-rate", 0))
        branch_rate = float(coverage.get("branch-rate", 0))

        print(f"📊 Line coverage: {line_rate:.2%}")
        print(f"📊 Branch coverage: {branch_rate:.2%}")

        # Check against minimum thresholds
        MIN_LINE_COVERAGE = 0.70  # 70%
        MIN_BRANCH_COVERAGE = 0.60  # 60%

        success = True

        if line_rate < MIN_LINE_COVERAGE:
            print(f"❌ Line coverage below threshold: {line_rate:.2%} < {MIN_LINE_COVERAGE:.2%}")
            success = False
        else:
            print(f"✅ Line coverage meets threshold: {line_rate:.2%} >= {MIN_LINE_COVERAGE:.2%}")

        if branch_rate < MIN_BRANCH_COVERAGE:
            print(f"❌ Branch coverage below threshold: {branch_rate:.2%} < {MIN_BRANCH_COVERAGE:.2%}")
            success = False
        else:
            print(f"✅ Branch coverage meets threshold: {branch_rate:.2%} >= {MIN_BRANCH_COVERAGE:.2%}")

        # Check per-module coverage
        packages = root.findall(".//package")
        for package in packages:
            package_name = package.get("name", "")
            classes = package.findall(".//class")

            for cls in classes:
                class_name = cls.get("name", "")
                module_line_rate = float(cls.get("line-rate", 0))
                full_name = f"{package_name}.{class_name}" if package_name else class_name

                if module_line_rate < MIN_LINE_COVERAGE * 0.5:  # 50% of minimum
                    print(f"⚠️  Low coverage in {full_name}: {module_line_rate:.2%}")

        return success

    except subprocess.CalledProcessError as e:
        print(f"❌ Error running coverage: {e}")
        return False
    except Exception as e:
        print(f"❌ Error parsing coverage: {e}")
        return False


if __name__ == "__main__":
    success = run_coverage_check()
    sys.exit(0 if success else 1)