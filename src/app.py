import logging
from src.config import settings
from src.tasks.process_lotes import run_process

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )
    run_process(settings)

if __name__ == "__main__":
    main()
