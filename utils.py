# Author : Sviatoslav Besnard pro@slavi.dev
# Created at: 2022-08-09

from typing import Optional, Union, Any, Callable

from rich.console import Group, OverflowMethod, JustifyMethod, Console
from rich.highlighter import Highlighter
from rich.progress import Progress, TaskID, SpinnerColumn, TextColumn, Task, ProgressColumn, BarColumn, \
    TaskProgressColumn, TimeRemainingColumn, MofNCompleteColumn
from rich.live import Live
from rich.style import Style, StyleType
from rich.table import Column
from rich.text import Text

class OneShotTaskContainer:
    """A class that contains a pending spinner task."""

    def __init__(self, task: TaskID, progress: Progress):
        self.task = task
        self.progress = progress
        self.ended = False

    def end(self) -> None:
        """End the task."""
        if not self.ended:
            self.progress.update(self.task, advance=1)
        self.ended = True


class RichConsole:
    """A singleton class that keep the progress for all the script"""
    live: Live = None

    """The verbose level of the console (0=error/warning, 1=general, 2=all)"""
    verbose: int = 2

    progress_bar: Progress = None
    progress_spinner: Progress = None

    @staticmethod
    def init() -> None:
        if RichConsole.live is None:
            RichConsole.progress_bar = Progress(
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                MofNCompleteColumn(),
                TimeRemainingColumn(),
            )
            RichConsole.progress_spinner = Progress(
                SpinnerColumn(),
                CallableTextColumn(
                    lambda task: f"{'  [progress.description]' if not task.finished else ':white_check_mark: '}"
                                 f"{task.description}"
                                 f"{'...' if not task.finished else ''}"),
            )

            RichConsole.live = \
                Live(
                    Group(
                        RichConsole.progress_bar,
                        RichConsole.progress_spinner),
                    console=Console(color_system="standard"))
            RichConsole.live.start()

    @staticmethod
    def close() -> None:
        if RichConsole.live is not None:
            RichConsole.live.stop()
            RichConsole.live = None

    @staticmethod
    def debug(msg: str) -> None:
        """Print a message if verbose level is set to 2"""
        if RichConsole.verbose >= 2:
            RichConsole.print(msg)

    @staticmethod
    def info(msg: str) -> None:
        """Print a message if verbose level is set to 1"""
        if RichConsole.verbose >= 1:
            RichConsole.print(msg)

    @staticmethod
    def warning(msg: str) -> None:
        """Print a message if verbose level is set to 0"""
        if RichConsole.verbose >= 0:
            RichConsole.print("[orange]Warning: " + msg)

    @staticmethod
    def error(msg: str) -> None:
        """Print a message if verbose level is set to 0"""
        if RichConsole.verbose >= 0:
            RichConsole.print("[red]ERROR: " + msg)

    @staticmethod
    def one_shot_task(title: str) -> OneShotTaskContainer:
        """Create a one-shot task"""
        return OneShotTaskContainer(task=RichConsole.progress_spinner.add_task(title, total=1, visible=True),
                                    progress=RichConsole.progress_spinner)

    @staticmethod
    def print(
            *objects: Any,
            sep: str = " ",
            end: str = "\n",
            style: Optional[Union[str, Style]] = None,
            justify: Optional[JustifyMethod] = None,
            overflow: Optional[OverflowMethod] = None,
            no_wrap: Optional[bool] = None,
            emoji: Optional[bool] = None,
            markup: Optional[bool] = None,
            highlight: Optional[bool] = None,
            width: Optional[int] = None,
            height: Optional[int] = None,
            crop: bool = True,
            soft_wrap: Optional[bool] = None,
            new_line_start: bool = False,
    ) -> None:
        """Print a message to the console"""
        RichConsole.live.console.print(
            *objects,
            sep=sep,
            end=end,
            style=style,
            justify=justify,
            overflow=overflow,
            no_wrap=no_wrap,
            emoji=emoji,
            markup=markup,
            highlight=highlight,
            width=width,
            height=height,
            crop=crop,
            soft_wrap=soft_wrap,
            new_line_start=new_line_start,
        )

    @staticmethod
    def print_author():
        """Print the author of the script"""
        RichConsole.print("[blue]Author:\n"
                          "[cyan]Name: [white]Sviatoslav Besnard\n"
                          "[cyan]Position: [white]Data Analyst Trainee\n"
                          "[cyan]Email: [white]pro@slavi.dev\n\n"
                          "[blue]Original idea and draft script:\n"
                          "[cyan]Name: [white]Hugo Grau\n"
                          "[cyan]Email: [white]hugo.grau@radissonhotels.com\n")


class CallableTextColumn(ProgressColumn):
    def __init__(
            self,
            text_format_callable: Callable[[Task], str],
            style: StyleType = "none",
            justify: JustifyMethod = "left",
            markup: bool = True,
            highlighter: Optional[Highlighter] = None,
            table_column: Optional[Column] = None,
    ) -> None:
        self.text_format_callable = text_format_callable
        self.justify: JustifyMethod = justify
        self.style = style
        self.markup = markup
        self.highlighter = highlighter
        super().__init__(table_column=table_column or Column(no_wrap=True))

    def render(self, task: "Task") -> Text:
        _text = self.text_format_callable(task)
        if self.markup:
            text = Text.from_markup(_text, style=self.style, justify=self.justify)
        else:
            text = Text(_text, style=self.style, justify=self.justify)
        if self.highlighter:
            self.highlighter.highlight(text)
        return text
