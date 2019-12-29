import configparser
import os

from atlassian import Confluence
from cryptography.fernet import Fernet
from jira import JIRA


class Connect:
    """Class to connect to JIRA and Confluence"""

    # Define Global variables
    global CurrentDir
    global config

    # Initiate Parameters
    CurrentDir = os.getcwd()

    # Parameters Ini
    config = configparser.ConfigParser()
    config.sections()
    config.read(CurrentDir + '\ini\Parameters.ini')

    def __init__(self):
        pass

    def Confluence(self):
        # Connect to confluence
        confluence = Confluence(
        url=config['CONNECT'].get('CONFserver'),
        username=config['CONNECT'].get('username'),
        password= self.decryptPass())

        return confluence;

    def Jira(self):
        # Connect to JIRA
        jira_server = config['CONNECT'].get('JIRAserver')
        jira_user = config['CONNECT'].get('username')
        jira_password = self.decryptPass()

        jira_server = {'server': jira_server}
        jira = JIRA(options=jira_server, basic_auth=(jira_user, jira_password))
        return jira;

    def GetConfig(self):
        return config;

    # Encrypt the password
    def CryptPass(self, msg):
        password = msg.encode()
        key = 'Xg-1FQOCobNN3dXpPojJpAR5rIb_d5ItQ4tWIAwZn28='
        f = Fernet(key)
        Passcrypt = f.encrypt(password)
        file = open(CurrentDir + '\ini\pykey.key', 'wb')
        file.write(Passcrypt)  # The key is type bytes still
        file.close()
        return

    # Decrypt the password
    def decryptPass(self):
        """ Method to Decrypt"""

        # Initiate Crypt Parameters
        key = 'Xg-1FQOCobNN3dXpPojJpAR5rIb_d5ItQ4tWIAwZn28='
        Kfile = '\ini\pykey.key'

        file = open(CurrentDir + Kfile, 'rb')
        Passdecrypt = file.read()  # The key will be type bytes
        file.close()
        f = Fernet(key)
        decrypted = f.decrypt(Passdecrypt)
        return (decrypted.decode("utf-8"))

