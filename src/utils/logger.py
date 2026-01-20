from loguru import logger
import sys

logger.remove()
logger.add(sys.stderr, level="INFO")
# Add file logging configuration here
