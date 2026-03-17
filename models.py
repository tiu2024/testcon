"""Data models for quiz questions and answer options."""

from dataclasses import dataclass, field


@dataclass
class Option:
    """A single answer option for a quiz question."""

    letter: str
    text: str
    is_correct: bool = False


@dataclass
class Question:
    """A multiple-choice quiz question with its answer options."""

    number: int
    text: str
    options: list[Option] = field(default_factory=list)

    @property
    def correct_option(self) -> Option | None:
        return next((o for o in self.options if o.is_correct), None)

    @property
    def incorrect_options(self) -> list[Option]:
        return [o for o in self.options if not o.is_correct]