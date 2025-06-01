import lib

if __name__ == "__main__":
    sheet = lib.get_data()
    soon = sheet.get_due_soon(7)
    overdue = sheet.get_overdue()
    message = ["Due soon:\n", soon, "\n\nOverdue:\n", overdue]
    lib.send_email(message)
