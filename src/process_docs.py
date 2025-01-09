import logging

import docx
import numpy as np
import pandas as pd

from src.data import Reports
from src.notification import TelegramAPI


def run(reports: Reports, end_date: str, bot: TelegramAPI) -> None:
    credit_columns = {
        "ID": "ID",
        "Дата начала": "start_date",
        "Номер договора": "contract_number",
    }

    credits_df = (
        pd.read_csv(
            reports.credit_contracts_fpath,
            sep="\t",
            skiprows=1,
            encoding="utf-16",
            usecols=list(credit_columns.keys()),
            engine="c",
        )
        .rename(columns=credit_columns)
        .dropna(axis=1, how="all")
        .dropna(subset=["ID"])
    )

    zbrk_l_deashd4_dirty_df = pd.read_excel(
        reports.zbrk_l_deashd4_xlsx_fpath, header=None, engine="calamine"
    )
    header_indices = zbrk_l_deashd4_dirty_df[
        zbrk_l_deashd4_dirty_df[1] == "Номер договора"
    ].index

    start_index = header_indices[0]
    if len(header_indices) > 1:
        end_index = header_indices[1]
    else:
        end_index = len(zbrk_l_deashd4_dirty_df)

    df = zbrk_l_deashd4_dirty_df.iloc[start_index:end_index]
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    df.rename(
        columns={
            "Клиент ": "client",
            "Дата погашения по графику": "deadline_date",
            "Валюта договора": "contract_currency",
            "Номер договора": "contract_number",
            "Проценты": "percentages",
            "Отсроченные проценты": "deferred_interest",
            "Основной долг": "debt",
        },
        inplace=True,
    )

    trash_start_idx = df[df["contract_number"] == "Всего"].index.min()
    df = df[0:trash_start_idx]

    df = df[df["deadline_date"] == end_date]

    for col in ["percentages", "deferred_interest", "debt"]:
        df.loc[:, col] = df[col].apply(
            lambda x: x if isinstance(x, float) else float(x.strip().replace(" ", ""))
        )

    for client in df.dropna(subset=["client"])["client"].unique():
        last_currency = ""

        deadline_dates = df[df["client"] == client]["deadline_date"].unique()
        for deadline_date in deadline_dates:
            client_name = client.replace('"', "")
            file_name = f"{client_name}_{end_date}.docx"
            doc_path = reports.docs_folder / file_name
            if doc_path.exists():
                logging.info(f"{file_name} exists. Skipping...")
                continue

            client_date_df: pd.DataFrame = df[
                (df["client"] == client) & (df["deadline_date"] == deadline_date)
            ]

            text = (
                f"\n\n\n\n\n\n{client}\n\n\n\nКасательно планового погашения по займу\n\n"
                f"Настоящим АО «Банк Развития Казахстана» сообщает, что {deadline_date} года наступает срок "
                f"погашения задолженности по следующим договорам банковского займа "
                f"заключенным между Банком и {client}."
                f"\nСумма к оплате:\n\n"
            )

            currencies = client_date_df["contract_currency"].unique()
            repayment_type1 = 0
            repayment_type2 = 0

            for currency in currencies:
                currency_client_date_df = client_date_df[
                    client_date_df["contract_currency"] == currency
                ]

                idx = 1
                total_sum = 0

                for row in currency_client_date_df.itertuples():
                    contract_currency = row.contract_currency
                    contract_number = row.contract_number
                    percentages = row.percentages
                    deferred_interest = row.deferred_interest
                    debt = row.debt

                    if contract_currency != last_currency:
                        text += f"{contract_currency}\n"
                        last_currency = contract_currency

                    credit_row = credits_df[
                        credits_df["contract_number"] == contract_number
                    ].iloc[0]
                    start_date = credit_row.start_date

                    text += f"\t{idx}. Вознаграждение по {contract_number} от {start_date} года - "

                    if np.isnan(row.deferred_interest):
                        percentages_str = f"{percentages:,.2f}".replace(",", " ")
                        text += f"{percentages_str} {currency}.\n"
                    else:
                        percentages += deferred_interest
                        full_percentages_str = f"{percentages:,.2f}".replace(",", " ")
                        text += f"{full_percentages_str} {currency}.\n"

                    idx += 1
                    repayment_type1 = 1

                    if not np.isnan(debt):
                        debt_str = f"{debt:,.2f}".replace(",", " ")
                        text += (
                            f"\t{idx}. Основной долг по {contract_number} от "
                            f"{start_date} года - {debt_str} {currency}.\n"
                        )

                        idx += 1
                        repayment_type2 = 1
                        total_sum += debt

                    total_sum += percentages

                total_sum = f"{total_sum:,.2f}".replace(",", " ")
                text += f"\nИтоговая сумма: {total_sum} {currency}.\n\n"

            if (repayment_type1 + repayment_type2) == 1:
                if repayment_type1 == 1:
                    repayment_text = "вознаграждения"
                else:
                    repayment_text = "основного долга"
            else:
                repayment_text = "вознаграждения и основного долга"

            text += (
                f"\nНа основании вышеизложенного просим Вас в срок "
                f"до {deadline_date} обеспечить в полном объёме денежные средства на "
                f"счете №KZ32907A287000000003, БИК DVKAKZKA в АО «Банк Развития Казахстана» "
                f"для планового погашения {repayment_text} по займам.\n\n"
                f"Надеемся на дальнейшее взаимовыгодное сотрудничество."
            )

            doc = docx.Document()
            doc.add_paragraph(text)
            doc.save(str(doc_path))
            logging.info(f'"{file_name}" saved...')

    bot.send_message("Documents are created...")
