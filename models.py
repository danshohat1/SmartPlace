
from dataclasses import dataclass, field
from typing import List, Union, Optional

@dataclass
class Student:
    name: str
    preferences: List[str]
    voice: float = 1.0
    match: Optional[str] = None

@dataclass
class University:
    name: str
    capacity: int
    preferences: List[Union[str, List[str]]]
    power: float = 1.0
    accepted: List[str] = field(default_factory=list)
    preferences_flat: List[str] = field(init=False)
    preference_pointer: int = field(default=0, init=False)

    def __post_init__(self):
        self.preferences_flat = []
        for tier in self.preferences:
            if isinstance(tier, list):
                self.preferences_flat.extend(tier)
            else:
                self.preferences_flat.append(tier)

    def has_free_slot(self):
        return len(self.accepted) < self.capacity
