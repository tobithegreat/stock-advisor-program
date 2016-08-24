import sys
import openpyxl
import datetime as dt

book = " "
def setup():
    """Here, the user will be prompted to enter name. Future implementation will have this function as a login system, so that you can access individual Excel worksheets. 
    """
    book = openpyxl.load_workbook('Investing/Stock_Advisor.xlsx')
    sheet = book.worksheets[0]
    name = raw_input("Hello! Please enter your name:  ").strip()
    validate_name(name, book)

    choice = raw_input("Hello {}. Please enter B to enter in a new stock buy, S to enter in a new stock sell, V to view your trade history, or Q to quit: ".format(name)).strip()
   
    if choice.upper() == "B":
        enter_new_buy(sheet)
    elif choice.upper == "S":
        enter_new_sale()
    elif choice.upper() == "V":
        view_trades()
    elif choice.upper() == "Q":
        sys.exit()


def validate_name(name,book):
    try:
        book.get_sheet_by_name(name)
    except KeyError:
        print("Sorry. You do not seem to have a worksheet set up currently with that name.\n")
        setup()


def enter_new_buy(sheet):
    stock = raw_input("What is the name of the stock?: ").strip()
    price_per_share = float(raw_input("How much is the price of each share?: ").strip())
    amount_of_shares = float(raw_input("How many shares are being purchased? ").strip())
    date_bought = dt.datetime.today().strftime("%m/%d/%Y %H:%M")
    gain_goal = raw_input(
"IMPORTANT: \n\n Your exit strategy is vital to being an efficient investor. What is your target gain percentage? Please enter as a whole number: ").strip()
    loss_goal = raw_input("What is your cut losses percentage? Please enter as a whole number: ").strip()
    total_cost = price_per_share * amount_of_shares
    
    print("CONFIRM: \n\n You have purchased {} share(s) of {}, at a price of {} for each share, for a total of {}. Your target gain is {}, and the loss in which you shall exit is {}.".format(amount_of_shares, stock, price_per_share, total_cost, gain_goal, loss_goal))

def view_trades(sheet):
    pass

if __name__ == "__main__":
    setup()
