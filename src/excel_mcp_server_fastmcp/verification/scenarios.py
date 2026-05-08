"""验证场景定义。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
TEST_DATA_DIR = ROOT / "tests" / "test_data"
VERIFICATION_TEST_DIR = ROOT / "tests" / "verification"
BASELINE_DIR = VERIFICATION_TEST_DIR / "baselines"
ARTIFACT_ROOT = ROOT / "artifacts" / "verification"


@dataclass(frozen=True)
class VerificationCase:
    """单个闭环验证用例。"""

    case_id: str
    kind: str
    fixture_path: Path
    sql: str
    sheet_name: str | None = None
    mutate: bool = False
    mutation_copy_name: str | None = None

    @property
    def baseline_path(self) -> Path:
        return BASELINE_DIR / f"{self.case_id}.json"

    @property
    def fixture_name(self) -> str:
        return self.fixture_path.name



def get_verification_cases() -> list[VerificationCase]:
    """返回当前项目的确定性验证场景。"""

    return [
        VerificationCase(
            case_id="select_game_config_top3",
            kind="select",
            fixture_path=TEST_DATA_DIR / "game_config.xlsx",
            sheet_name="技能配置",
            sql="SELECT skill_id, skill_name, damage, cooldown FROM 技能配置 ORDER BY skill_id LIMIT 3",
        ),
        VerificationCase(
            case_id="select_join_preview",
            kind="select",
            fixture_path=TEST_DATA_DIR / "join_test.xlsx",
            sheet_name="技能表",
            sql="SELECT * FROM 技能表 LIMIT 3",
        ),
        VerificationCase(
            case_id="update_game_config_damage",
            kind="update",
            fixture_path=TEST_DATA_DIR / "game_config.xlsx",
            sheet_name="技能配置",
            sql="UPDATE 技能配置 SET damage = damage + 1 WHERE skill_id = 'SK001'",
            mutate=True,
            mutation_copy_name="game_config_mutation.xlsx",
        ),
    ]
