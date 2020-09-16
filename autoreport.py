import time
import win32com.client
import sys
import csv
import os
import re
import pyodbc
from pyad import aduser
# spróbować adquery
from tabulate import tabulate
from pyad.pyadexceptions import invalidResults
from display_data import display_data

# Zmienne globalne dajemy z dużej litery
# wyjtek, django_urls
# bardziej podejscie klasowe

outlook = win32com.client.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")

# CHEIDS PASSWORD VAULT'S INBOX MESSAGES
o = win32com.client.Dispatch("Outlook.Application")
folder = outlook.Folders("Password Vault")
inbox = folder.Folders("Inbox")
messages = inbox.Items

# Script sometimes doesn't get newest messages and this command solves the problem
messages.Sort("[ReceivedTime]", True)

# Keywords -> We will later cheID if the message subject has any of them
keywords = ["SOME KEYWORDS"]


# Regex patterns to get info names from e-mail message
regex_patterns = ["""SOME REGEX PATTERNS"""]


def find_info_in_access_db(info_numbers: list) -> tuple:
    """
    Finds owners and delegate info_owners and adds their IDs
        to info_owners list (later will be used to find usernames by their IDs in AD)
    Adds info's info to info list i.e. (Column1, Column2, Column3)
        as one list element (will be useful to print the data)

    :return: info and info_owners lists
    """
    info = []
    info_owners = []

    def check_ID_if_info_owner_in_row(IDs: list) -> list:
        """
        CheIDs if info_owner or delegate info_owner are provided
        :param IDs:  info_owners and/or delegate info_owners IDs
        :return: a list with IDs converted to uppercase
        """
        ID_exist = []
        for ID in IDs:
            try:
                ID_exist.append(ID.upper())
            except AttributeError:
                continue
        return ID_exist

    if info_numbers:
        conn = pyodbc.connect(
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="""PATH TO ACCESS DB"""')
        cursor = conn.cursor()
        cursor.execute(
            'select FullName, OwnerID, DelegateID from info')
        for row in cursor.fetchall():
            for info in info_numbers:
                if row.FullName.startswith(info):
                    IDs_found_in_row = check_ID_if_info_owner_in_row(
                        row[1:3])
                    for ID in IDs_found_in_row:
                        if ID not in info_owners:
                            info_owners.append(ID)
                    info.append(row)
    return info, info_owners


def get_user_info_From_ad(info_owners: list) -> dict:
    """
    Gets usernames by given IDs list from Active Directory and puts them into so_do_dict
    :param info_owners: a list with owners and delegates IDs
    :return: a dictionary, where IDs are keys() and usernames are values()
    """
    so_do_dict = {}
    for ID in info_owners:
        try:
            user = aduser.ADUser.from_cn(ID)
            wrong_names = ["Wrong Name 1",
                           'Wrong Name 2']

            user_wrong_ad_name = bool(
                user.Description in wrong_names)

            so_do_dict[ID] = user.Description if not user_wrong_ad_name else user.DisplayName

        except invalidResults:
            continue
    return so_do_dict


def get_info_owners_emails(info_owners: list) -> dict:
    """
    Finds owners and delegates e-mails and puts
    them into info_owners_emails dict.
    :param info_owners: a list of owners and delegates IDs
    """
    info_owners_emails = {}
    for ID in info_owners:
        user = aduser.ADUser.from_cn(ID)
        info_owners_emails[ID] = user.UserPrincipalName
    return info_owners_emails


def find_info_in_message_body_by_regex(regex_patterns: list, message: object) -> list:
    """
    Finds info names by given regex patterns
    :param regex_patterns: a list with regex patterns to use while getting info names from message body
    :param message: a message object being currently analyzed
    :return: message_body_info_found list with info names found in message body
    """
    message_body_info_found = []
    for regex in regex_patterns:
        pattern = re.compile(regex)
        for info in pattern.findall(message.Body):
            if info not in message_body_info_found:
                message_body_info_found.append(info)
    return message_body_info_found


def create_data_to_display(info: list, so_do_dict: dict) -> list:
    """
    Creates a list, with lists as elements, which will be then used to display to the end user with tabulate library
    :param info: a list with tuple elements in format ("info name", "owner ID", "delegate ID)
    :param so_do_dict: a dictionary with owners and delegate IDs as keys() and their usernames as values()
        so_do_dict stands for info_owners and delegate owners dictionary
    :return: a info_with_info_owners list to display later with tabulate library
    """
    info_with_info_owners = []
    for info_data in info:
        info_with_info_owners.append(
            [info_data[0], so_do_dict.get(info_data[1]), so_do_dict.get(info_data[2])])
    return info_with_info_owners


def find_approval(message_body: str, sender_name: str) -> dict:
    """
    Finds the approvals of the potential info_owners or delegate info_owners.
    Takes the entire body of the message and cuts it on sub messages with regex ("From") keyword.
    :return: approvals found in the conversation
    """

    def add_approval_with_date_to_approvals(approver: str, approvals: dict, message_date: str) -> dict:
        """
        Adds the approver as key and approval date as value to approvals dictionary.
        If the key already exists, simply return current dict. This is to get the newest approvals in the message.
        :return: updated approvals dictionary
        """
        try:
            approvals[approver]
            return approvals
        except KeyError:
            approvals[approver] = message_date
            return approvals

    def clear_message(message):
        """
        Clears the message from useless characters
        :return: cleared message
        """
        elems_to_clear = ["\r", "\n", "*\t", '"']
        for elem in elems_to_clear:
            message = message.replace(elem, "")
        message = re.sub(" +", " ", message)
        return message

    def find_approval_date_with_regex(message: str) -> str:
        """
        Finds message date
        :param message: a message in conversation
        :return: message's date
        """
        time_received_regexes = [
            r"\d{1,2} \w+ \d{4} \d{1,2}:\d{1,2}", r"\w+ \d{1,2}, \d{4} \d{1,2}:\d{1,2}", r"\d{1,2} .*? \d{4} at \d{1,2}:\d{1,2}"]

        message_date = False
        for regex in time_received_regexes:
            r = re.compile(regex)
            if r.findall(message):
                message_date = r.findall(message)[0]
                break
        return message_date

    def get_approvals(sender_data: list, approvals: dict, message: str, message_date: str):
        """
        Gets the approvals from a message
        :param sender_data: a list that contains one element with sender data (list, because the data was captured with regex.findall method)
        :param approvals: approvals found in the message so far
        :param message_date: date of the message sent
        :return: updated approvals dictionary
        """
        def add_approval_with_date_to_approvals(approver: str, approvals: dict, message_date: str) -> dict:
            """
            Adds the approver as key and approval date as value to approvals dictionary.
            If the key already exists, simply return current dict. This is to get the newest approvals in the message.
            :param approver: a person that included approval keyword in a message
            :param approvals: as above
            :param message_date: as above
            :return: as above
            """
            try:
                approvals[approver]
                return approvals
            except KeyError:
                approvals[approver] = message_date
                return approvals

        message_date = message_date if sender_data else "Inbox"
        try:
            sender_data = sender_data_regex.findall(
                message)[0][2:-2]
        except IndexError:
            sender_data = sender_name

        approvals = add_approval_with_date_to_approvals(
            sender_data, approvals, message_date)
        return approvals

    cleared_message = clear_message(message_body)
    approval_keywords = ["approval", "approve", "approved"]
    conversation = cleared_message.split("From")
    sender_data_regex = re.compile(r"^: .*? <")
    approvals = {}

    for message in conversation:
        if any(approval_keyword in message.lower() for approval_keyword in approval_keywords):
            message_date = find_approval_date_with_regex(
                cleared_message)
            sender_data = sender_data_regex.findall(message)
            approvals = get_approvals(
                sender_data, approvals, cleared_message, message_date)
    for key in list(approvals):
        if key.startswith('"'):
            try:
                approvals[key] = approvals.pop(key).replace('"', '')
            except:
                ...
    return approvals


def job():
    """
    For each message with a keyword in keywords list in subject:
    1. Decode message subject (prevent errors coming from non-ascii characters)
    2. Get sender e-mail address if the sender is not Password Vault
    3. Parse the message body and find info with regex
    4. Get info data from Access Database
    5. Get info_owners and delegate info_owners data from Active Directory
    6. Find approvals in message body
    7. Create data to display with tabulate library
    """
    print("START")
    for message in messages:
        message_subject_decoded = message.subject.encode(
            "ascii", "ignore").decode()
        if any(keyword in message_subject_decoded.lower() for keyword in keywords):
            sender_email = [
                message.Sender.GetExchangeUser().PrimarySmtpAddress]
            sender_name = message.Sender.GetExchangeUser().Name
            if sender_email[0] == """ ENTER YOUR INBOX EMAIL""":
                continue
            else:
                message_body_info_found = find_info_in_message_body_by_regex(
                    regex_patterns, message)
                info, info_owners = find_info_in_access_db(
                    message_body_info_found)
                so_do_dict = get_user_info_From_ad(
                    info_owners)
                info_owners_emails = get_info_owners_emails(
                    info_owners)
                message_body_encoded = message.Body.encode(
                    "ascii", "ignore").decode()
                body_emails = re.compile(
                    r'[\w.]*@[\w.]*').findall(message_body_encoded)
                approvals = find_approval(
                    message.Body, sender_name)
                info_with_info_owners = create_data_to_display(
                    info, so_do_dict)
                display_data(
                    message, sender_email[0], approvals, info_with_info_owners, info)


job()
