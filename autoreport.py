import time
import win32com.client
import sys
import csv
import os
import re
import pyodbc
from pyad import aduser
from tabulate import tabulate
from pyad.pyadexceptions import invalidResults
from display_data import display_data


outlook = win32com.client.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")

# CHECKS PASSWORD VAULT'S INBOX MESSAGES
o = win32com.client.Dispatch("Outlook.Application")
folder = outlook.Folders("Password Vault")
inbox = folder.Folders("Inbox")
messages = inbox.Items

# Script sometimes doesn't get newest messages and this command solves the problem
messages.Sort("[ReceivedTime]", True)

# Keywords -> We will later check if the message subject has any of them
keywords = ["report", "reports", "inventory", "activity",
            "activities", "raport", "testowanie skryptu",
            "extract", "scs", "soll", "ist", "npa review", "usage"]


# Regex patterns to get safe names from e-mail message
regex_patterns = [r'\d{5}', r'PAAPPL[\w_-]*', r'PAPPL[\w_-]*',
                  r'DTAPAPPLU[\w_-]*', r'DTAPINFR[\w_-]*', r'DTAPPLU[\w_-]*', r'TAPAPPL[\w_-]*', r'DAPAPPLU[\w_-]*']


def find_safe_owners_and_safe_data(safe_numbers: list) -> tuple:
    """
    Finds owners and delegate safe owners and adds their CKs
        to owners_delegates list (later will be used to find usernames by their CKs in AD)
    Adds safe's info to safes_data list i.e. (FullName, OwnerCK, DelegateCK)
        as one list element (will be useful to print the data)

    :return: safes_data and owners_delegates lists
    """
    safes_data = []
    owners_delegates = []

    def check_if_owner_or_delegate_CK_in_row(CKs: list) -> list:
        """
        Checks if safe owner or delegate safe owner are provided
        :param CKs:  safe owners and/or delegate safe owners CKs
        :return: a list with CKs converted to uppercase
        """
        CK_exist = []
        for CK in CKs:
            try:
                CK_exist.append(CK.upper())
            except AttributeError:
                continue
        return CK_exist

    if safe_numbers:
        conn = pyodbc.connect(
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=Z:\MCDB_PRD.accdb;')
        cursor = conn.cursor()
        cursor.execute(
            'select FullName, OwnerCK, DelegateCK from safes')
        for row in cursor.fetchall():
            for safe in safe_numbers:
                if row.FullName.startswith(safe):
                    CKs_found_in_row = check_if_owner_or_delegate_CK_in_row(
                        row[1:3])
                    for CK in CKs_found_in_row:
                        if CK not in owners_delegates:
                            owners_delegates.append(CK)
                    safes_data.append(row)
    return safes_data, owners_delegates


def get_owners_and_delegates_from_ad(owners_delegates: list) -> dict:
    """
    Gets usernames by given CKs list from Active Directory and puts them into so_do_dict
    :param owners_delegates: a list with owners and delegates CKs
    :return: a dictionary, where CKs are keys() and usernames are values()
    """
    so_do_dict = {}
    for ck in owners_delegates:
        try:
            user = aduser.ADUser.from_cn(ck)
            try:
                so_do_dict[ck] = user.Description
            except UnicodeEncodeError:
                so_do_dict[ck] = str(
                    user.Description, encoding='utf-8', errors='ignore')
        except invalidResults:
            continue
    return so_do_dict


def get_owners_delegates_emails(owners_delegates: list) -> dict:
    """
    Finds owners and delegates e-mails and puts
    them into owners_delegates_emails dict.
    :param owners_delegates: a list of owners and delegates CKs
    """
    owners_delegates_emails = {}
    for ck in owners_delegates:
        user = aduser.ADUser.from_cn(ck)
        owners_delegates_emails[ck] = user.UserPrincipalName
    return owners_delegates_emails


def find_safes_in_message_body_by_regex(regex_patterns: list, message: object) -> list:
    """
    Finds safe names by given regex patterns
    :param regex_patterns: a list with regex patterns to use while getting safe names from message body
    :param message: a message object being currently analyzed
    :return: message_body_safes_found list with safe names found in message body
    """
    message_body_safes_found = []
    for regex in regex_patterns:
        pattern = re.compile(regex)
        for safe in pattern.findall(message.Body):
            if safe not in message_body_safes_found:
                message_body_safes_found.append(safe)
    return message_body_safes_found


def create_safe_data_to_display(safes_data: list, so_do_dict: dict) -> list:
    """
    Creates a list, with lists as elements, which will be then used to display to the end user with tabulate library
    :param safes_data: a list with tuple elements in format ("safe name", "owner CK", "delegate CK)
    :param so_do_dict: a dictionary with owners and delegate CKs as keys() and their usernames as values()
        so_do_dict stands for safe owners and delegate owners dictionary
    :return: a safes_with_owners_delegates list to display later with tabulate library
    """
    safes_with_owners_delegates = []
    for safe_data in safes_data:
        safes_with_owners_delegates.append(
            [safe_data[0], so_do_dict.get(safe_data[1]), so_do_dict.get(safe_data[2])])
    return safes_with_owners_delegates


def find_approval(message_body: str, sender_name: str) -> dict:
    """
    Finds the approvals of the potential safe owners or delegate safe owners.
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
    3. Parse the message body and find safes with regex
    4. Get safes data from Access Database
    5. Get safe owners and delegate safe owners data from Active Directory
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
            if sender_email[0] == "password.vault@ing.com":
                continue
            else:
                message_body_safes_found = find_safes_in_message_body_by_regex(
                    regex_patterns, message)
                safes_data, owners_delegates = find_safe_owners_and_safe_data(
                    message_body_safes_found)
                so_do_dict = get_owners_and_delegates_from_ad(
                    owners_delegates)
                owners_delegates_emails = get_owners_delegates_emails(
                    owners_delegates)
                message_body_encoded = message.Body.encode(
                    "ascii", "ignore").decode()
                body_emails = re.compile(
                    r'[\w.]*@[\w.]*').findall(message_body_encoded)
                approvals = find_approval(
                    message.Body, sender_name)
                safes_with_owners_delegates = create_safe_data_to_display(
                    safes_data, so_do_dict)
                display_data(
                    message, sender_email[0], approvals, safes_with_owners_delegates, safes_data)


job()
