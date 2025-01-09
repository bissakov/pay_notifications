import logging
import sys
import warnings
from pathlib import Path

import dotenv
from urllib3.exceptions import InsecureRequestWarning

project_folder = Path(__file__).resolve().parent.parent
sys.path.append(str(project_folder))

from src.notification import TelegramAPI
from src import robot
from src.utils import logger
from src.utils.colvir import kill_all_processes

if __name__ == "__main__":
    logger.setup_logger(project_folder)

    if sys.version_info.major != 3 or sys.version_info.minor != 12:
        logging.error(f"Python {sys.version_info} is not supported")
        raise RuntimeError(f"Python {sys.version_info} is not supported")

    warnings.simplefilter(action="ignore", category=UserWarning)
    warnings.simplefilter(action="ignore", category=InsecureRequestWarning)
    env_path = project_folder / ".env"
    dotenv.load_dotenv(env_path)

    kill_all_processes("COLVIR")
    kill_all_processes("EXCEL")
    kill_all_processes("WINWORD")

    telegram_bot = TelegramAPI()
    robot.run(bot=telegram_bot, project_folder=project_folder, env_path=env_path)
