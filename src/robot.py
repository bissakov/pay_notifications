import logging
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from time import sleep

import pywinauto.timings
from dotenv import set_key

from src import process_docs
from src.data import Date, TimeRange, Reports
from src.notification import TelegramAPI, handle_error, Mail, send_mail
from src.utils.colvir import ColvirInfo, Colvir
from src.utils.excel_utils import is_file_exported, convert_report, Excel


def get_from_env(key: str) -> str:
    value = os.getenv(key)
    assert isinstance(value, str), f"{key} not set in the environment variables"
    return value


def backup_env_file(now: datetime, env_path: Path, backup_folder: Path) -> None:
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    backup_path = backup_folder / f".env.backup_{timestamp}"
    shutil.copy(env_path, backup_path)
    logging.info(f"Backup created: {backup_path}")


def export_files(
    reports: Reports,
    t_range: TimeRange,
    bot: TelegramAPI,
    env_path: Path,
    backup_folder: Path,
) -> None:
    credit_contracts_exists = reports.credit_contracts_fpath.exists()
    zbrk_l_deashd4_exists = reports.zbrk_l_deashd4_fpath.exists()
    zbrk_l_deashd4_xlsx_exists = reports.zbrk_l_deashd4_xlsx_fpath.exists()

    if credit_contracts_exists and zbrk_l_deashd4_xlsx_exists:
        msg = "No need in exporting. Retuning early..."
        logging.info(msg)
        bot.send_message(msg)
        return

    if (
        credit_contracts_exists
        and zbrk_l_deashd4_exists
        and not zbrk_l_deashd4_xlsx_exists
    ):
        msg = (
            f"{reports.credit_contracts_fpath} exists but "
            f"{reports.zbrk_l_deashd4_xlsx_fpath} is not yeat converted..."
        )
        logging.info(msg)
        bot.send_message(msg)
        with Excel() as excel:
            convert_report(
                excel=excel,
                source=reports.zbrk_l_deashd4_fpath,
                dist=reports.zbrk_l_deashd4_xlsx_fpath,
            )
        bot.send_message(f"{reports.zbrk_l_deashd4_xlsx_fpath} converted...")
        return

    colvir_info = ColvirInfo(
        loader=Path(get_from_env("LOADER_PATH")),
        colvir=Path(get_from_env("COLVIR_PATH")),
        user=get_from_env("COLVIR_USER"),
        password=get_from_env("COLVIR_PASSWORD"),
    )
    logging.info(f"{colvir_info=}")

    with Colvir(colvir_info=colvir_info, bot=bot) as colvir:
        if colvir.was_password_changed:
            backup_env_file(t_range.start.dt, env_path, backup_folder)
            set_key(env_path, "COLVIR_PASSWORD", colvir.info.password)

        colvir.choose_mode("SLOAN")

        filter_win = colvir.utils.get_window(title="Фильтр")
        filter_win.wait(wait_for="enabled")
        filter_win["OK"].click()

        credits_win = colvir.utils.get_window(title="Кредитные договора")
        if not credit_contracts_exists:
            msg = f"{reports.credit_contracts_fpath.name} does not exist. Exporting..."
            logging.info(msg)
            bot.send_message(msg)

            for i in range(5):
                credits_win.menu_select("#4->#4->#1")
                colvir.save_excel(file_path=reports.credit_contracts_fpath)
                if (error_win := colvir.app.window(title="Произошла ошибка")).exists():
                    error_msg = error_win.child_window(class_name="Edit").window_text()
                    logging.warning(f"{error_msg=}")
                    error_win.close()
                if not reports.credit_contracts_fpath.exists():
                    continue
                else:
                    break
            else:
                raise Exception("Unable to export credit_contracts")

        if not zbrk_l_deashd4_xlsx_exists:
            if zbrk_l_deashd4_exists:
                msg = f"{reports.zbrk_l_deashd4_fpath.name} exists. Converting..."
                logging.info(msg)
                bot.send_message(msg)

                with Excel() as excel:
                    convert_report(
                        excel=excel,
                        source=reports.zbrk_l_deashd4_fpath,
                        dist=reports.zbrk_l_deashd4_xlsx_fpath,
                    )
            else:
                msg = f"{reports.zbrk_l_deashd4_xlsx_fpath.name} does not exist. Exporting and converting..."
                logging.info(msg)
                bot.send_message(msg)

                try:
                    credits_win.wait(wait_for="exists enabled", timeout=20)
                except pywinauto.timings.TimeoutError:
                    colvir.reload()
                    colvir.choose_mode("SLOAN")

                    filter_win = colvir.utils.get_window(title="Фильтр")
                    filter_win["OK"].click()
                    credits_win = colvir.utils.get_window(title="Кредитные договора")

                if not credits_win.has_focus():
                    credits_win.set_focus()

                colvir.find_and_click_button(
                    window=credits_win,
                    toolbar=credits_win["Static0"],
                    target_button_name="Получить отчет(F5)",
                )

                report_win = colvir.utils.get_window(title="Выбор отчета")
                colvir.utils.click_input(report_win["Предварительный просмотр"])
                colvir.utils.click_input(report_win["Экспорт в файл..."])
                file_win = colvir.utils.get_window(title="Файл отчета ")

                colvir.utils.type_keys(
                    file_win["Edit2"],
                    str(reports.report_root_folder),
                    step_delay=0.3,
                    delay_after=1,
                )
                colvir.utils.type_keys(
                    file_win["Edit4"], reports.zbrk_l_deashd4_fpath.name, step_delay=0.3
                )
                try:
                    file_win["ComboBox"].select(12)
                    sleep(1)
                except (IndexError, ValueError):
                    pass
                file_win["OK"].click()

                params_win = colvir.utils.get_window(title="Параметры отчета ")
                params_win["Edit2"].set_text(t_range.start.short)
                params_win["Edit4"].set_text(t_range.end.short)
                params_win["OK"].click()

                with Excel() as excel:
                    status = False
                    while not status:
                        sleep(5)
                        message, status = is_file_exported(
                            file_path=reports.zbrk_l_deashd4_fpath, excel=excel
                        )
                        logging.info(message)

                    if not reports.zbrk_l_deashd4_xlsx_fpath.exists():
                        logging.info(
                            f"'{reports.zbrk_l_deashd4_xlsx_fpath.name}' does not exist. Converting original..."
                        )
                        convert_report(
                            excel=excel,
                            source=reports.zbrk_l_deashd4_fpath,
                            dist=reports.zbrk_l_deashd4_xlsx_fpath,
                        )


@handle_error
def run(bot: TelegramAPI, project_folder: Path, env_path: Path) -> None:
    today_dt = datetime.now()
    # today_dt = datetime(2024, 12, 20)
    current_year_month_name = today_dt.strftime("%Y_%m/%d.%m.%y")

    t_range = TimeRange(
        start=Date.to_date(today_dt), end=Date.to_date(today_dt + timedelta(days=16))
    )

    logging.info(f"{t_range=}")

    backup_folder = project_folder / "backups"
    backup_folder.mkdir(exist_ok=True)
    logging.info(f"{backup_folder=}")

    report_root_folder = project_folder / "reports" / current_year_month_name
    report_root_folder.mkdir(parents=True, exist_ok=True)
    logging.info(f"{report_root_folder=}")

    docs_folder = report_root_folder / "docs"
    docs_folder.mkdir(exist_ok=True)
    logging.info(f"{docs_folder=}")

    bot.send_message(
        f"Старт процесса за {t_range.start.short}\n" f'"Уведомления по план плате"'
    )

    credit_contracts_fpath = report_root_folder / f"credits_{t_range.start.short}.xls"
    zbrk_l_deashd4_fpath = (
        report_root_folder / f"ZBRK_L_DEASHD4_{t_range.start.short}.xls"
    )
    zbrk_l_deashd4_xlsx_fpath = zbrk_l_deashd4_fpath.with_suffix(".xlsx")

    reports = Reports(
        report_root_folder=report_root_folder,
        docs_folder=docs_folder,
        credit_contracts_fpath=credit_contracts_fpath,
        zbrk_l_deashd4_fpath=zbrk_l_deashd4_fpath,
        zbrk_l_deashd4_xlsx_fpath=zbrk_l_deashd4_xlsx_fpath,
    )

    export_files(
        reports=reports,
        t_range=t_range,
        bot=bot,
        backup_folder=backup_folder,
        env_path=env_path,
    )

    process_docs.run(reports=reports, end_date=t_range.end.short, bot=bot)

    mail_info = Mail(
        server=os.getenv("SMTP_SERVER"),
        sender=os.getenv("SMTP_SENDER"),
        recipients=os.getenv("SMTP_RECIPIENTS"),
        subject="Отчет робота по плановым платежам",
        attachment_folder_path=reports.docs_folder,
    )

    send_mail(mail_info=mail_info, t_range=t_range, bot=bot)

    bot.send_message("Успешное окончание процесса")
    logging.info("Successfully finished...")
