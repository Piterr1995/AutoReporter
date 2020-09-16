from tabulate import tabulate


def display_data(message: object, sender_email: str, approvals: dict, info: list, data: list):
    print("\n====================================================================================",
          f"\n{message.Subject} \nFrom: {sender_email}\n\nApprovals:")

    for key, value in approvals.items():
        print(key, ": ", value)

    print(
        "------------------------------------------------------------------------------------")

    print(tabulate(info, headers=[
        "Info1", "Info2", 'Info3']))

    print(
        "------------------------------------------------------------------------------------\n\nINFO FOUND: ")
    for i in info:
        print(i[0])

    print(
        "====================================================================================\n")
