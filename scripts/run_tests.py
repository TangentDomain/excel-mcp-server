# Test runner script
import subprocess
import sys
from pathlib import Path

def run_tests():
    """Run all tests and generate coverage report"""
    project_root = Path(__file__).parent
    
    # Install coverage if not already installed
    try:
        import coverage
    except ImportError:
        print("Installing coverage...")
        subprocess.run([sys.executable, "-m", "pip", "install", "coverage"], 
                      cwd=project_root, check=True)
    
    # Run tests with coverage
    print("Running tests with coverage...")
    cmd = [
        sys.executable, "-m", "coverage", "run", 
        "-m", "pytest", 
        "tests/", 
        "-v", 
        "--tb=short",
        "--cov=src",
        "--cov-report=html",
        "--cov-report=term-missing",
        "--cov-fail-under=80"
    ]
    
    result = subprocess.run(cmd, cwd=project_root)
    
    if result.returncode == 0:
        print("\n✅ All tests passed!")
        print("📊 Coverage report generated in htmlcov/ directory")
    else:
        print("\n❌ Some tests failed!")
        return result.returncode
    
    return 0

if __name__ == "__main__":
    sys.exit(run_tests())