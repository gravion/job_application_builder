#!/usr/bin/env python
import logging
import re
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml
from jinja2 import Environment, FileSystemLoader

# --- 0. setting up  ---
# logging

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("Main")


# path definitions
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
ATTACHMENT_DIR = BASE_DIR / "attachments"
CONFIG_FILE = BASE_DIR / "config.yaml"
APPLICATION_DIR = BASE_DIR / "applications"
HISTORY_FILE = BASE_DIR / "application_history.csv"
CONTENT_FILE = BASE_DIR / "application_text.txt"
# variables
ERROR_SYMBOL = "❌"
CHECKMARK_SYMBOL = "✓"
WARNING_SYMBOL = "⚠"
COGWHEEL_SYMBOL = "⚙"
CROSS_SYMBOL = "✗"
# --- 2. environment functions ---


def get_jinja_env():
    """
    Setting up Jinja2 for LaTeX environment.
    especially resolving bracket confusion.
    """

    def regex_replace(s, find, replace):
        return re.sub(find, replace, s)

    env = Environment(
        loader=FileSystemLoader(str(TEMPLATE_DIR)),
        # redefine delimiter
        block_start_string="<%",
        block_end_string="%>",
        variable_start_string="<<",
        variable_end_string=">>",
        comment_start_string="<#",
        comment_end_string="#>",
        trim_blocks=True,
        lstrip_blocks=True,
    )
    env.filters["regex_replace"] = regex_replace

    return env


def create_shell_friendly_name(text):
    """
    Removes non letter and nonnumbersfrom string.
    """
    text = text.replace(" ", "_").replace("/", "-")

    return re.sub(r"[^a-zA-Z0-9_-]", "", text)


def load_config():
    if not CONFIG_FILE.exists():
        logger.error(f"{CROSS_SYMBOL} {CONFIG_FILE} not found!")

        return {}
    with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def render_document(env, template_name, output_name, data):
    """
    Rendering a template and moving it to the template folder.
    """
    try:
        template = env.get_template(template_name)
        output_path = TEMPLATE_DIR / output_name

        with open(output_path, "w", encoding="utf-8") as fh:
            fh.write(template.render(data))
        logger.info(f"{CHECKMARK_SYMBOL} {output_name} generated.")
    except Exception as err:
        logger.error(f"{CROSS_SYMBOL} Error in rendering of {template_name}: {err}")


def compile_latex(tex_file):
    """
    Runnig xelatex for the template folder
    """
    logger.info(f"{COGWHEEL_SYMBOL} Compiling {tex_file} ...")

    program = "xelatex"
    option1 = "-interaction=nonstopmode"
    option2 = "-halt-on-error"

    try:
        subprocess.run(
            [program, option1, option2, tex_file],
            cwd=TEMPLATE_DIR,
            check=True,
            capture_output=True,
            text=True,
        )
        logger.info(f"{CHECKMARK_SYMBOL} Pdf creation successful.")
    except subprocess.CalledProcessError as err:
        error_msg = err.stderr if err.stderr else "Unknown LaTeX Error : {err}"
        last_lines = "\n".join(error_msg.splitlines()[-20:])
        logger.error(f"{CROSS_SYMBOL} Error in compiling {tex_file}:\n{last_lines}")


def archive_results(data):
    """
    Creating taget folder.
    Moves files and rename.
    """
    # --- creating structure for filenames
    data_str = datetime.now().strftime("%Y-%m-%d")
    company = create_shell_friendly_name(data.get("company", "unknown_company")).lower()
    job = create_shell_friendly_name(data.get("job_title", "job")).lower()

    # --- creating target folder if not exists
    folder_name = f"{data_str}_{company}_{job}"
    target_dir = APPLICATION_DIR / folder_name
    target_dir.mkdir(parents=True, exist_ok=True)

    # --- moving files and renaming
    main_src = TEMPLATE_DIR / "main.pdf"

    for name in ["cv.pdf", "application_letter.pdf", "attachments.pdf"]:
        shutil.move(str(TEMPLATE_DIR / name), target_dir / name)
    shutil.copy(str(CONFIG_FILE), target_dir / "config.yaml")
    shutil.copy(
        str(BASE_DIR / "application_letter_text.txt"),
        target_dir / "application_letter_text.txt",
    )

    if main_src.exists():
        final_name = f"application_{company}_{job}.pdf"
        shutil.move(str(main_src), str(target_dir / final_name))
        logger.info(f"{CHECKMARK_SYMBOL} main pdf moved to: {final_name}.")


def log_to_history(data):
    """
    Inserts appliction into history file, if line does not exist.
    """

    if not HISTORY_FILE.exists():
        logger.warning(f"⚠ {HISTORY_FILE} file not found. Skipping.")

        return

    date = datetime.now().strftime("%Y-%m-%d")
    week = datetime.now().isocalendar()[1]
    company = data.get("company", "Unknown")
    jobtitle = data.get("job_title", "Unknown")

    try:
        df = pd.read_csv(HISTORY_FILE)
        # Check ob heute schon für diese Firma geloggt wurde
        exists = (
            (df["date"] == date)
            & (df["company"] == company)
            & (df["position"] == jobtitle)
        ).any()

        if not exists:
            new_row = [date, week, company, jobtitle]
            df.loc[len(df)] = new_row
            df.to_csv(HISTORY_FILE, index=False)
            logger.info(f"✓ Entry added to {HISTORY_FILE}.")
        else:
            logger.info(f"ℹ Entry of {company} already exists.")
    except Exception as e:
        logger.error(f"✗  Error while logging: {e}")


def main():
    # loading data

    data = load_config()

    if not data:
        logger.error(f"{CROSS_SYMBOL} Error loading Data.")

        return

    lang = data.get("language", "de")
    lang_data = data.get(lang, {})
    data.update(lang_data)
    data["lang"] = lang

    if lang == "en":
        data["doc_lang"] = "english"
    elif lang == "de":
        data["doc_lang"] = "ngerman"
        # add other languages here
    else:
        logger.critical(
            f"{CROSS_SYMBOL} No valid language set: {lang}. Only de/en is available"
        )

    # init jinja env
    env = get_jinja_env()

    # validating attachments
    valid_attachments = []

    for doc_name in data.get("attachments", []):
        doc_path = ATTACHMENT_DIR / doc_name

        if doc_path.exists():

            valid_attachments.append(doc_name)
        else:
            logger.warning(f"{WARNING_SYMBOL} Attachment is missing {doc_path}")
    data["attachments"] = valid_attachments

    if CONTENT_FILE.exists():
        with open(CONTENT_FILE, "r", encoding="utf-8") as f:
            data["content"] = f.read()
    else:
        data["content"] = "Missing application letter text."
        logger.warning(
            f"{WARNING_SYMBOL} application_letter_text.txt content hasn't bee found!"
        )
    # rendering and compiling documents

    docs_to_build = [
        ("attachments.tex.j2", "attachments.tex"),
        ("cv.tex.j2", "cv.tex"),
        ("application_letter.tex.j2", "application_letter.tex"),
        ("main.tex.j2", "main.tex"),
    ]

    for template, output in docs_to_build:
        render_document(env, template, output, data)
        compile_latex(output)

    logger.info("Rendering and compiling done.")
    archive_results(data)
    log_to_history(data)


if __name__ == "__main__":
    main()
