import dataclasses
import logging
import os
import random
import re
import secrets
import string
from pathlib import Path
from time import sleep
from types import TracebackType
from typing import Type

import psutil
import pyautogui
import pyperclip
import pywinauto
import pywinauto.base_wrapper
import pywinauto.timings
import win32con
import win32gui
from attr import define
from pywinauto import mouse, win32functions

from src.notification import TelegramAPI

pyautogui.FAILSAFE = False


@dataclasses.dataclass(slots=True)
class ColvirInfo:
    loader: Path
    colvir: Path
    user: str
    password: str


@define
class DialogContent:
    title: str | None
    content: str | None
    button_names: list[str]

    def __getitem__(self, item: str) -> str | list[str] | None:
        return getattr(self, item)

    def __setitem__(self, key: str, value: (str | None) | list[str]) -> None:
        setattr(self, key, value)


def kill_all_processes(proc_name: str) -> None:
    for proc in psutil.process_iter():
        try:
            if proc_name in proc.name():
                proc.terminate()
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue


def generate_password(
    length: int = 12,
    min_digits: int = 1,
    min_low_letters: int = 1,
    min_up_letters: int = 1,
    min_punctuations: int = 1,
) -> str:
    def random_chars(char_set: str, min_count: int) -> str:
        return "".join(secrets.choice(char_set) for _ in range(min_count))

    digits = random_chars(string.digits, min_digits)
    low_letters = random_chars(string.ascii_lowercase, min_low_letters)
    up_letters = random_chars(string.ascii_uppercase, min_up_letters)
    punctuations = random_chars(string.punctuation, min_punctuations)

    required_length = min_digits + min_low_letters + min_up_letters + min_punctuations
    remaining_length = max(0, length - required_length)
    all_characters = string.ascii_letters + string.digits
    remaining_chars = random_chars(all_characters, remaining_length)

    password_chars = list(
        digits + low_letters + up_letters + punctuations + remaining_chars
    )
    secrets.SystemRandom().shuffle(password_chars)

    password = "".join(password_chars)
    return password


class ColvirUtils:
    def __init__(self, app: pywinauto.Application | None) -> None:
        self.app = app

    @staticmethod
    def wiggle_mouse(duration: int) -> None:
        def get_random_coords() -> tuple[int, int]:
            screen = pyautogui.size()
            width = screen[0]
            height = screen[1]

            return random.randint(100, width - 200), random.randint(100, height - 200)

        max_wiggles = random.randint(4, 9)
        step_sleep = duration / max_wiggles

        for _ in range(1, max_wiggles):
            coords = get_random_coords()
            pyautogui.moveTo(x=coords[0], y=coords[1], duration=step_sleep)

    @staticmethod
    def close_window(
        win: pywinauto.WindowSpecification, raise_error: bool = False
    ) -> None:
        if win.exists():
            win.close()
            return

        if raise_error:
            raise pywinauto.findwindows.ElementNotFoundError(
                f"Window {win} does not exist"
            )

    @staticmethod
    def set_focus_win32(win: pywinauto.WindowSpecification) -> None:
        if win.wrapper_object().has_focus():
            return

        handle = win.wrapper_object().handle

        mouse.move(coords=(-10000, 500))
        if win.is_minimized():
            if win.was_maximized():
                win.maximize()
            else:
                win.restore()
        else:
            win32gui.ShowWindow(handle, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(handle)

        win32functions.WaitGuiThreadIdle(handle)

    @staticmethod
    def set_focus(win: pywinauto.WindowSpecification, retries: int = 20) -> None:
        while retries > 0:
            try:
                if retries % 2 == 0:
                    ColvirUtils.set_focus_win32(win)
                else:
                    win.set_focus()
                break
            except (Exception, BaseException):
                retries -= 1
                sleep(5)
                continue

        if retries <= 0:
            raise Exception("Failed to set focus")

    @staticmethod
    def press(win: pywinauto.WindowSpecification, key: str, pause: float = 0) -> None:
        ColvirUtils.set_focus(win)
        win.type_keys(key, pause=pause, set_foreground=False)

    @staticmethod
    def type_keys(
        window: pywinauto.WindowSpecification,
        keystrokes: str,
        step_delay: float = 0.1,
        delay_after: float = 0.5,
    ) -> None:
        ColvirUtils.set_focus(window)
        for command in list(filter(None, re.split(r"({.+?})", keystrokes))):
            try:
                window.type_keys(command, set_foreground=False)
            except pywinauto.base_wrapper.ElementNotEnabled:
                sleep(1)
                window.type_keys(command, set_foreground=False)
            sleep(step_delay)

        sleep(delay_after)

    def get_window(
        self,
        title: str,
        wait_for: str = "exists",
        timeout: int = 20,
        regex: bool = False,
        found_index: int = 0,
    ) -> pywinauto.WindowSpecification:
        if regex:
            window = self.app.window(title_re=title, found_index=found_index)
        else:
            window = self.app.window(title=title, found_index=found_index)
        window.wait(wait_for=wait_for, timeout=timeout)
        sleep(0.5)
        return window

    def persistent_win_exists(self, title_re: str, timeout: float) -> bool:
        try:
            self.app.window(title_re=title_re).wait(wait_for="enabled", timeout=timeout)
        except pywinauto.timings.TimeoutError:
            return False
        return True

    def close_dialog(self) -> None:
        dialog_win = self.get_window(title="Colvir Banking System", found_index=0)
        dialog_win.set_focus()
        sleep(0.5)
        dialog_win["OK"].click_input()

    @staticmethod
    def click_input(window: pywinauto.WindowSpecification) -> None:
        if not window.has_focus():
            window.set_focus()
        sleep(0.5)
        window.click_input()


class Colvir:
    def __init__(self, colvir_info: ColvirInfo, bot: TelegramAPI) -> None:
        kill_all_processes(proc_name="AppLoader")
        kill_all_processes(proc_name="COLVIR")
        self.info = colvir_info
        self.app: pywinauto.Application | None = None
        self.utils = ColvirUtils(app=self.app)
        self.bot = bot
        self.was_password_changed = False

    def open_colvir(self) -> None:
        original_dir = os.getcwd()
        apploader_dir = self.info.loader.parent
        os.chdir(apploader_dir)

        if os.getenv("ENV") == "prod":
            start_app_location = self.info.loader
        else:
            start_app_location = self.info.colvir

        for _ in range(10):
            try:
                app_loader = pywinauto.Application().start(
                    cmd_line=str(start_app_location)
                )
                sleep(2)

                if (dialog := app_loader.Dialog).exists():
                    dialog["OK"].click_input()

                self.app = pywinauto.Application().connect(path=str(self.info.colvir))
                self.login()
                self.check_interactivity()
                break
            except (Exception, BaseException):
                kill_all_processes("AppLoader")
                kill_all_processes("COLVIR")
                continue
        assert self.app is not None, Exception("max_retries exceeded")
        self.utils.app = self.app

        os.chdir(original_dir)

        os.chdir(original_dir)

    def login(self) -> None:
        login_win = self.app.window(title="Вход в систему")

        login_username = login_win["Edit2"]
        login_password = login_win["Edit"]

        login_username.set_text(text=self.info.user)
        if login_username.window_text() != self.info.user:
            login_username.set_text("")
            login_username.type_keys(self.info.user, set_foreground=False)

        login_password.set_text(text=self.info.password)
        if login_password.window_text() != self.info.password:
            login_password.set_text("")
            login_password.type_keys(self.info.password, set_foreground=False)

        login_win["OK"].click()

        sleep(1)
        if (
            login_win.exists()
            and (error_win := self.app.window(title="Произошла ошибка")).exists()
        ):
            error_msg = error_win.child_window(class_name="Edit").window_text()
            logging.error(error_msg)
            raise pywinauto.findwindows.ElementNotFoundError()

    def check_interactivity(self) -> None:
        if isinstance(self.app, pywinauto.Application):
            self.utils.app = self.app

        if (attention_win := self.app.window(title="Внимание")).exists():
            attention_win.close()
            change_pass_win = self.utils.get_window(title="Смена пароля.+", regex=True)

            new_password = generate_password(
                min_digits=3, min_low_letters=3, min_up_letters=3, min_punctuations=0
            )
            logging.info(f"Not yet changed - {new_password=}")

            for i in range(3):
                change_pass_win["Edit2"].set_text(new_password)
                change_pass_win["Edit0"].set_text(new_password)
                ok_button = change_pass_win["OK"]
                if ok_button.is_enabled() is False:
                    continue
                else:
                    ok_button.click()
                    sleep(1)
                    break
            else:
                raise RuntimeError("Unable to change the password")

            if attention_win.exists():
                raise ValueError(
                    "Password probably violates contraints - min 3 digits and min 3 letters"
                )

            self.was_password_changed = True
            self.info.password = new_password
            logging.info(f"Successfully changed - {new_password=}")

        self.choose_mode(mode="TREPRT")
        sleep(1)

        reports_win = self.app.window(title="Выбор отчета")
        self.utils.close_window(win=reports_win, raise_error=True)

    def choose_mode(self, mode: str) -> None:
        mode_win = self.app.window(title="Выбор режима")
        mode_win["Edit2"].set_text(text=mode)
        self.utils.press(mode_win["Edit2"], "~")

    @staticmethod
    def parse_dialog_content(dialog_text: str) -> DialogContent:
        lines = list(filter(lambda l: l, dialog_text.split("\r\n")))

        dialog_content = DialogContent(title=None, content=None, button_names=[])

        section = None
        for line in lines:
            if line.startswith("[Window Title]"):
                section = "title"
            elif line.startswith("[Content]"):
                section = "content"
            elif (
                line.startswith("[OK]")
                or line.startswith("[Cancel]")
                or line.startswith("[")
            ):
                dialog_content.button_names.append(line.strip("[]"))
            else:
                if section:
                    dialog_content[section] = line
                    section = None

        return dialog_content

    def dialog_text(self) -> str | None:
        dialog_win = self.app.window(title="Colvir Banking System", found_index=0)
        if not dialog_win.exists():
            return None

        if not dialog_win.has_focus():
            dialog_win.set_focus()
            sleep(0.5)

        dialog_win.type_keys("^C")
        sleep(0.5)
        dialog_text = pyperclip.paste()

        dialog_content = self.parse_dialog_content(dialog_text=dialog_text)
        dialog_content_text = dialog_content.content
        if dialog_content_text is not None:
            dialog_win.close()
        return dialog_content_text

    def find_and_click_button(
        self,
        window: pywinauto.WindowSpecification,
        toolbar: pywinauto.WindowSpecification,
        target_button_name: str,
        horizontal: bool = True,
        offset: int = 5,
    ) -> tuple[int, int]:
        if not window.has_focus():
            window.set_focus()

        status_win = self.app.window(title_re="Банковская система.+")
        rectangle = toolbar.rectangle()
        mid_point = rectangle.mid_point()
        mouse.move(coords=(mid_point.x, mid_point.y))

        start_point = rectangle.left if horizontal else rectangle.top
        end_point = mid_point.x if horizontal else mid_point.y

        x, y = mid_point.x, mid_point.y
        point = 0

        x_offset = offset if horizontal else 0
        y_offset = offset if not horizontal else 0

        error_count = 0

        i = 0
        while (
            active_button := status_win["StatusBar"].window_text().strip()
        ) != target_button_name:
            if point > end_point:
                logging.error(f"{point=}, {end_point=}")
                logging.error(f"{active_button=}, f{target_button_name=}")

                point = 0
                i = 0
                error_count += 1

                if error_count >= 3:
                    raise pywinauto.findwindows.ElementNotFoundError
                continue

            point = start_point + i * 5

            if horizontal:
                x = point
            else:
                y = point

            mouse.move(coords=(x, y))
            i += 1

        x += x_offset
        y += y_offset

        mouse.click(button="left", coords=(x, y))

        return x, y

    def save_excel(self, file_path: Path) -> None:
        assert file_path.name.endswith(".xls"), "Only .xls files are supported"
        kill_all_processes("EXCEL")

        file_win = self.utils.get_window(title="Выберите файл для экспорта")

        if file_path.exists():
            file_path.unlink()

        file_win["Edit4"].set_text(str(file_path))
        file_win["&Save"].click_input()

        sleep(1)

        sort_win = self.utils.get_window(title="Сортировка")
        sort_win["OK"].click()

        sleep(5)
        while not file_path.exists():
            sleep(10)

        kill_all_processes("EXCEL")

    def reload(self) -> None:
        self.exit()
        self.open_colvir()

    def exit(self) -> None:
        if not self.app.kill():
            kill_all_processes("COLVIR")

    def __enter__(self) -> "Colvir":
        self.open_colvir()
        return self

    def __exit__(
        self,
        exc_type: Type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ):
        if exc_val is not None or exc_type is not None or exc_tb is not None:
            self.bot.send_image()
        self.exit()
