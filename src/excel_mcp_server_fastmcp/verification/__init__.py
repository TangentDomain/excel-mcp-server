"""验证子包：基于固定 fixture + baseline 的确定性闭环验证。"""

from .runner import run_verification
from .scenarios import VerificationCase, get_verification_cases

__all__ = ["VerificationCase", "get_verification_cases", "run_verification"]
