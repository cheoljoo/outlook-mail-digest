# -*- coding: utf-8 -*-
"""
OneDrive에 동기화된 OutlookDigest.xlsx 를 원격 서버로 SSH 키 인증을 통해 전송합니다.

- SSH 키가 없으면 생성하고, 서버에 등록되어 있지 않으면 등록합니다 (이미 등록되어 있으면 skip).
- Windows 작업 스케줄러에 이 스크립트를 주기적으로 실행하는 작업이 없으면 스스로 등록합니다
  (이미 등록되어 있으면 skip).
- 마지막으로 파일을 복사합니다.
- 모든 단계는 시간과 함께 sync_outlook_digest.log 에 기록되어, 나중에 실패 원인이
  (OneDrive 파일 없음 / 서버 연결 실패 등) 무엇이었는지 로그만 보고 알 수 있습니다.

run.bat(같은 폴더)를 통해 `uv run sync_outlook_digest.py`로 실행되는 것을 전제로 합니다.
Windows의 내장 OpenSSH 클라이언트(ssh-keygen/ssh/scp)와 schtasks만 사용하므로 별도 pip 설치가
필요 없습니다.
"""

import logging
import os
import subprocess
import sys
from pathlib import Path

REMOTE_USER = "cheoljoo.lee"
REMOTE_HOST = "tiger02.lge.com"
REMOTE_DIR = "~/temp/"
KEY_PATH = Path.home() / ".ssh" / "id_rsa"
LOCAL_FILE = Path(os.environ.get("OneDriveCommercial", "")) / "OutlookDigest.xlsx"
TASK_NAME = "OutlookDigestSync"
RUN_BAT_PATH = Path(__file__).resolve().parent / "run.bat"
LOG_PATH = Path(__file__).resolve().parent / "sync_outlook_digest.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("sync_outlook_digest")


def ensure_ssh_key() -> None:
    pub_key_path = KEY_PATH.with_suffix(".pub")
    if KEY_PATH.exists() and pub_key_path.exists():
        logger.info(f"SSH 키가 이미 있습니다: {KEY_PATH}")
        return

    logger.info("SSH 키가 없어 새로 생성합니다...")
    KEY_PATH.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        ["ssh-keygen", "-t", "rsa", "-b", "4096", "-N", "", "-f", str(KEY_PATH)],
        check=True,
    )
    logger.info(f"SSH 키 생성 완료: {KEY_PATH}")


def is_key_registered() -> bool:
    result = subprocess.run(
        [
            "ssh",
            "-o", "BatchMode=yes",
            "-o", "ConnectTimeout=5",
            "-i", str(KEY_PATH),
            f"{REMOTE_USER}@{REMOTE_HOST}",
            "echo ok",
        ],
        capture_output=True,
        text=True,
    )
    return result.returncode == 0 and "ok" in result.stdout


def register_key() -> None:
    pub_key = KEY_PATH.with_suffix(".pub").read_text().strip()
    logger.info("서버에 SSH 키를 등록합니다. 최초 1회 비밀번호 입력이 필요합니다.")

    remote_cmd = (
        "mkdir -p ~/.ssh && chmod 700 ~/.ssh && "
        f"grep -qxF '{pub_key}' ~/.ssh/authorized_keys 2>/dev/null || "
        f"echo '{pub_key}' >> ~/.ssh/authorized_keys && "
        "chmod 600 ~/.ssh/authorized_keys"
    )
    try:
        subprocess.run(
            ["ssh", f"{REMOTE_USER}@{REMOTE_HOST}", remote_cmd],
            check=True,
        )
    except subprocess.CalledProcessError as e:
        logger.error(f"SSH 키 등록 실패 (서버 연결 불가/비밀번호 오류 가능): {e}")
        sys.exit(1)

    if not is_key_registered():
        logger.error("키 등록 후에도 비밀번호 없이 접속이 안 됩니다. 등록 결과를 확인해주세요.")
        sys.exit(1)
    logger.info("SSH 키 등록 완료.")


def is_task_registered() -> bool:
    result = subprocess.run(
        ["schtasks", "/Query", "/TN", TASK_NAME],
        capture_output=True,
        text=True,
    )
    return result.returncode == 0


def register_task() -> None:
    logger.info(f"작업 스케줄러에 '{TASK_NAME}'가 없어 새로 등록합니다 (매시간 반복 실행: {RUN_BAT_PATH}).")
    try:
        subprocess.run(
            [
                "schtasks", "/Create",
                "/TN", TASK_NAME,
                "/TR", f'"{RUN_BAT_PATH}"',
                "/SC", "HOURLY",
                "/MO", "1",
                "/RL", "LIMITED",
                "/F",
            ],
            check=True,
        )
    except subprocess.CalledProcessError as e:
        logger.error(f"작업 스케줄러 등록 실패: {e}")
        sys.exit(1)
    logger.info(f"작업 스케줄러 등록 완료: {TASK_NAME}")


def copy_file() -> None:
    if not LOCAL_FILE.exists():
        logger.error(
            f"로컬 파일을 찾을 수 없습니다: {LOCAL_FILE} "
            "(OneDrive 동기화가 안 됐거나, 파일이 삭제/이동됐거나, OneDriveCommercial 환경변수가 잘못됐을 가능성)"
        )
        sys.exit(1)

    logger.info(f"전송 시작: {LOCAL_FILE} -> {REMOTE_USER}@{REMOTE_HOST}:{REMOTE_DIR}")
    result = subprocess.run(
        [
            "scp",
            "-o", "ConnectTimeout=10",
            "-i", str(KEY_PATH),
            str(LOCAL_FILE),
            f"{REMOTE_USER}@{REMOTE_HOST}:{REMOTE_DIR}",
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        logger.error(
            f"파일 전송 실패 (서버 다운/네트워크 문제 가능, 종료 코드 {result.returncode}): "
            f"{result.stderr.strip()}"
        )
        sys.exit(1)
    logger.info("파일 전송 성공.")


def main() -> None:
    logger.info("===== sync_outlook_digest 시작 =====")
    try:
        ensure_ssh_key()
        if is_key_registered():
            logger.info("SSH 키가 이미 서버에 등록되어 있습니다. 재등록하지 않습니다.")
        else:
            register_key()

        if is_task_registered():
            logger.info(f"작업 스케줄러에 '{TASK_NAME}'가 이미 등록되어 있습니다. 재등록하지 않습니다.")
        else:
            register_task()

        copy_file()
    except SystemExit:
        logger.error("===== sync_outlook_digest 실패로 종료 =====")
        raise
    except Exception:
        logger.exception("예상하지 못한 오류로 실패")
        logger.error("===== sync_outlook_digest 실패로 종료 =====")
        sys.exit(1)
    else:
        logger.info("===== sync_outlook_digest 성공적으로 종료 =====")


if __name__ == "__main__":
    main()
