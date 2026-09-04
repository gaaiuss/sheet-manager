import os

from dotenv import load_dotenv
from rich import print

load_dotenv()

greetings = os.getenv("GREETINGS", "`.env` file not found, check `.env-example`")

print(greetings)
