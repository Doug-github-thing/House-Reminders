import lib

if __name__ == "__main__":
    sheet = lib.get_data()
    soon = sheet.get_due_soon(7)
    overdue = sheet.get_overdue()
    message = ["Due soon:", soon, "Overdue:", overdue]

    html = lib.format_email_html(message)
    lib.send_email(html)
