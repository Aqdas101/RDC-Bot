from loguru import logger

logger.remove()

logger.add("../logs/success.log", level="INFO", filter=lambda record: record["extra"].get("status") == "success")
logger.add("../logs/failure.log", level="ERROR", filter=lambda record: record["extra"].get("status") == "failure")

def log_success(message):
    logger.bind(status="success").info(message)

def log_failure(message):
    logger.bind(status="failure").error(message)


