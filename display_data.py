from tabulate import tabulate


def display_data(message: object, sender_email: str, approvals: dict, safes_with_owners_delegates: list, safes_data: list):
    print("\n====================================================================================",
          f"\n{message.Subject} \nFrom: {sender_email}\n\nApprovals:")

    for key, value in approvals.items():
        print(key, ": ", value)

    print(
        "------------------------------------------------------------------------------------")

    print(tabulate(safes_with_owners_delegates, headers=[
        "Safe name", "Safe owner", 'Delegate safe owner']))

    print(
        "------------------------------------------------------------------------------------\n\nSAFES FOUND: ")
    for safe in safes_data:
        print(safe[0])

    print(
        "====================================================================================\n")
