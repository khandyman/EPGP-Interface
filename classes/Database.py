from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from os import environ


class Database:
    """
    This class provides all database interaction
    """

    def __init__(self):
        load_dotenv()

        # obtain database parameters securely as environment variables
        self._MYSQL_USER = environ.get('MYSQL_USER')
        self._MYSQL_PASSWORD = environ.get('MYSQL_PASSWORD')
        self._MYSQL_HOST = environ.get('MYSQL_HOST')
        self._MYSQL_DB = environ.get('MYSQL_DB')

        # set connection string
        self._params = f"mysql+pymysql://{self._MYSQL_USER}:{self._MYSQL_PASSWORD}@{self._MYSQL_HOST}/{self._MYSQL_DB}?charset=utf8mb4"

    ################# READ METHODS #################
    def find_main(self, char_name):
        """
        Get main character for a given discord id
        :discord_id: discord id to look up
        :return: results of the select query, in list form
        """
        query = (
            "SELECT char_name, char_race, char_class from sos_bot.characters WHERE "
            f"discord_id = (SELECT discord_id FROM sos_bot.characters WHERE char_name = "
            f"'{char_name}') AND char_type = 'Main'"
        )

        return self.execute_read(query)

    ################# UTILITY METHODS #################
    def create_engine(self):
        """
        create a new engine object upon request, to prevent
        MySQL server from timing out
        :return: SQL alchemy engine object
        """
        # create the query engine
        return create_engine(self._params)

    def execute_read(self, query):
        """
        Send a read query to database engine
        :query: the formatted query string to send
        to database engine
        :return: results of the operation, in list form
        """
        records_list = []

        # open a connection to database, dynamically close
        # it when with block closes
        with self.create_engine().connect() as conn:
            result = conn.execute(text(query))

            # get query results and, line by line,
            # convert to dict entries; add each
            # dict to list
            for row in result.all():
                row_to_dict = row._asdict()
                records_list.append(row_to_dict)

            conn.close()

        return records_list
