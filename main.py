from dotenv import load_dotenv
import os
import requests
import json
from helper import test

load_dotenv()

if __name__ == "__main__":
    print(test.TestAPI())
