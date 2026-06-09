from __future__ import annotations

import random
import time


def sleep_range(min_seconds: float, max_seconds: float) -> float:
    delay = random.uniform(min_seconds, max_seconds)
    time.sleep(delay)
    return delay


def backoff(attempt: int, base: float = 1.5, jitter: float = 0.3) -> float:
    delay = (base ** max(0, attempt - 1)) + random.uniform(0, jitter)
    time.sleep(delay)
    return delay

