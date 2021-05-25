import os

from dotenv import load_dotenv

from modules import update_statistics


load_dotenv('.env')

if not os.environ.get('TOKEN'):
    raise EnvironmentError("TOKEN is not provided")


if __name__ == '__main__':
    update_statistics()
