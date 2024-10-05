# tratamento_tabela_xlsx_html_xls
Este script tem como finalidade enviar um e-mail diariamente, com uma tabela tratada com as especificações necessárias para suspensão de contas bancárias. 

# bibliotecas
import pandas as pd <br>
from datetime import datetime <br>
import os <br>
import smtplib <br>
from email.mime.multipart import MIMEMultipart <br>
from email.mime.text import MIMEText <br>
from email.mime.base import MIMEBase <br>
from email import encoders <br>
import shutil
