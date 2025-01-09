import email.utils
import io
import logging
import os
import shutil
import smtplib
import traceback
import urllib.parse
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from functools import wraps
from pathlib import Path
from typing import NamedTuple, Callable

import requests
import requests.adapters
from requests.exceptions import SSLError
from typing import cast
import PIL.Image as Image
import PIL.ImageGrab as ImageGrab

from src.data import TimeRange


def get_secrets() -> tuple[str, str]:
    token = os.getenv("TOKEN")
    if token is None:
        raise EnvironmentError('Environment variable "TOKEN" is not set.')

    chat_id = os.getenv("CHAT_ID")
    if chat_id is None:
        raise EnvironmentError('Environment variable "CHAT_ID" is not set.')

    return token, chat_id


class TelegramAPI:
    def __init__(self) -> None:
        self.session = requests.Session()
        self.session.mount("http://", requests.adapters.HTTPAdapter(max_retries=5))
        self.token, self.chat_id = get_secrets()
        self.api_url = f"https://api.telegram.org/bot{self.token}/"

    def reload_session(self) -> None:
        self.session = requests.Session()
        self.session.mount("http://", requests.adapters.HTTPAdapter(max_retries=5))

    def send_message(
        self, message: str, use_session: bool = True, use_md: bool = False
    ) -> bool:
        send_data: dict[str, str | None] = {
            "chat_id": self.chat_id,
        }

        if use_md:
            send_data["parse_mode"] = "MarkdownV2"

        files = None

        url = urllib.parse.urljoin(self.api_url, "sendMessage")
        send_data["text"] = message

        try:
            if use_session:
                response = self.session.post(
                    url, data=send_data, files=files, verify=False
                )
            else:
                response = requests.post(url, data=send_data, files=files, verify=False)

            data = "" if not hasattr(response, "json") else response.json()
            logging.info(f"{response.status_code=}")
            logging.info(f"{data=}")
            response.raise_for_status()
            return response.status_code == 200
        except SSLError as err:
            logging.exception(err)
            return False

    def send_image(
        self, media: Image.Image | None = None, use_session: bool = True
    ) -> bool:
        try:
            send_data = {"chat_id": self.chat_id}

            url = urllib.parse.urljoin(self.api_url, "sendPhoto")

            image_stream = io.BytesIO()
            if media is None:
                media = ImageGrab.grab()
            media.save(image_stream, format="PNG")
            image_stream.seek(0)
            raw_io_base_stream = cast(io.RawIOBase, image_stream)
            buffered_reader = io.BufferedReader(raw_io_base_stream)
            files = {"photo": buffered_reader}

            if use_session:
                response = self.session.post(url, data=send_data, files=files)
            else:
                response = requests.post(url, data=send_data, files=files)

            data = "" if not hasattr(response, "json") else response.json()
            logging.info(f"{response.status_code=}")
            logging.info(f"{data=}")
            response.raise_for_status()
            return response.status_code == 200
        except requests.exceptions.ConnectionError as exc:
            logging.exception(exc)
            return False

    def send_with_retry(
        self,
        message: str,
    ) -> bool:
        retry = 0
        while retry < 5:
            try:
                use_session = retry < 5
                success = self.send_message(message, use_session)
                return success
            except (
                requests.exceptions.ConnectionError,
                requests.exceptions.SSLError,
                requests.exceptions.HTTPError,
            ) as e:
                self.reload_session()
                logging.exception(e)
                logging.warning(f"{e} intercepted. Retry {retry + 1}/10")
                retry += 1

        return False


def handle_error(func: Callable[..., any]) -> Callable[..., any]:
    @wraps(func)
    def wrapper(*args, **kwargs) -> any:
        bot: TelegramAPI | None = kwargs.get("bot")

        try:
            return func(*args, **kwargs)
        except KeyboardInterrupt as error:
            raise error
        except (Exception, BaseException) as error:
            logging.exception(error)
            error_msg = traceback.format_exc()

            developer = os.getenv("DEVELOPER")
            if developer:
                error_msg = f"@{developer} {error_msg}"

            if bot:
                bot.send_message(error_msg)
            raise error

    return wrapper


class Mail(NamedTuple):
    server: str
    sender: str
    recipients: str
    subject: str
    attachment_folder_path: Path


def is_folder_empty(folder_path: Path) -> bool:
    for _ in folder_path.iterdir():
        return False
    return True


def send_mail(mail_info: Mail, t_range: TimeRange, bot: TelegramAPI) -> bool:
    recipients_lst: list[str] = mail_info.recipients.split(";")

    msg = MIMEMultipart()
    msg["From"] = mail_info.sender
    msg["To"] = mail_info.recipients
    msg["Date"] = email.utils.formatdate(localtime=True)
    msg["Subject"] = mail_info.subject

    body = mail_info.subject

    if is_folder_empty(mail_info.attachment_folder_path):
        body += f"\n\nНа {t_range.end.short} г. нет плановых платежей по займам."
        bot.send_message("No documents")
    else:
        bot.send_message(
            f"{len(os.listdir(mail_info.attachment_folder_path))} new documents"
        )

        archive_name = f"Documents_{t_range.end.short}"
        original_zip_path = shutil.make_archive(
            base_name=archive_name,
            format="zip",
            root_dir=mail_info.attachment_folder_path,
        )
        archive_name = f"{archive_name}.zip"
        doc_archive_path = mail_info.attachment_folder_path / archive_name
        shutil.move(original_zip_path, doc_archive_path)

        with open(doc_archive_path, "rb") as f:
            part = MIMEApplication(f.read())
        part.add_header("Content-Disposition", "attachment", filename=archive_name)
        msg.attach(part)

    msg.attach(MIMEText(body, "html", "utf-8"))

    try:
        with smtplib.SMTP(mail_info.server, 25) as smtp:
            response = smtp.sendmail(mail_info.sender, recipients_lst, msg.as_string())
            if response:
                logging.error("Failed to send email to the following recipients:")
                for recipient, error in response.items():
                    logging.error(f"{recipient}: {error}")
                bot.send_message("Email sent unsuccessfully...")
                return False
            else:
                logging.info("Email sent successfully...")
                bot.send_message("Email sent successfully...")
                return True
    except smtplib.SMTPException as e:
        logging.error(f"Failed to send email: {e}")
        bot.send_message("Email sent unsuccessfully...")
        return False
