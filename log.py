import pathlib
import datetime


class Log:
    def add_new_error(self, error):
        with open(f"{pathlib.Path().resolve()}\\data\\log.txt", "a") as myfile:
            myfile.write(f"\n{datetime.datetime.now()}-{error}")