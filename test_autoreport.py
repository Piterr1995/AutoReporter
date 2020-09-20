import pytest
import autoreport
from email_messages import email_message_with_info, email_message_with_approval, second_email_message_with_approval


@pytest.mark.slow
def test_find_info_owners_and_data():
    data, owners_delegates = autoreport.find_info_owners_and_data(['111111', '00000', '2222222',
                                                                              "3333333", "44444444444444",
                                                                              "55555555", "666666666"])
    assert str(safes_data) == str([
        ('1111111111', 'XX11XX', 'ZZ11ZZ'),
        ('2222222222', 'XX11XX', 'ZZ11ZZ'),
        ('3333333331', 'XX11XX', 'ZZ11ZZ'),
        ('4444444444', 'XX11XX', 'ZZ11ZZ'),
        ('5555555555', 'XX11XX', 'ZZ11ZZ'),
        ('6666666666', 'XX11XX', 'ZZ11ZZ'),
    ])
    assert owners_delegates == ['XX11XX', "ZZ11ZZ"]


@pytest.mark.fast
def test_get_owners_and_delegates_from_ad():
    so_do_dict = autoreport.get_owners_and_delegates_from_ad(
        ['JJ68JJ', 'xxxxxxxx', 'CC60CC'])
    no_data_so_do_dict = autoreport.get_owners_and_delegates_from_ad([
    ])
    assert so_do_dict == {
        'JJ68JJ': 'John Doe', "CC60CC": 'John Doe'}
    assert no_data_so_do_dict == {}


@pytest.mark.fast
def test_get_owners_delegates_emails():
    owners_delegates_emails = autoreport.get_owners_delegates_emails(
        ['AA12AA', 'AA12AA',  'AA12AA', 'AA12AA'])
    assert owners_delegates_emails == {'AA12AA': 'xxxxxxx@xxx.com', 'AA12AA': 'zzzzzzzz@zzz.com',
                                       'AA12AA': 'aaaaaaaaa@aaa.com', 'AA12AA': 'bbbbbbbbbbb@bbbbbbb.com'}


@pytest.mark.fast
def test_find_info_in_message_body_by_regex():
    regex_patterns = ["""Some regex patterns"""]
    email_string = email_message_with_info()

    class Message():
        def __init__(self, Body):
            self.Body = Body
    message = Message(Body=email_string)
    message_body_info_found = autoreport.find_info_in_message_body_by_regex(
        regex_patterns, message)
    assert message_body_info_found == ['000000', '11111111', '222222']


@pytest.mark.fast
def test_find_approval():
    approval_keywords = ["approve", "approved", "approval"]
    conversation = mail_message_with_approval()
    sender_name = 'John Doe'
    assert autoreport.find_approval(conversation, sender_name) == {
        'John Doe': 'Inbox', 'John Doe': '6 augustus 2020 21:41', 'John Doe': 'August 6, 2020 5:35'}
