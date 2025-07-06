import logging
import pandas as pd
import re

# 导入必要的提取器函数
from .case_extractors_names import (
    extract_name_from_case_report,
    extract_name_from_decision,
    extract_name_from_trial_report
)


from .case_extractors_demographics import (
    extract_suspected_violation_from_case_report,
    extract_suspected_violation_from_decision
)

logger = logging.getLogger(__name__)


# 简要案情验证规则已移动到 case_validation_additional.py 中

