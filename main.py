import pandas as pd
from datetime import datetime


def Main():
    # Ask user for location of bank activity file
    BankAct = input("What is the location of the bank activity file?:")
    Mapping = "./Mapping/Cash_Rec_Mapping.xlsx"

    # Load bank and mapping files into dataframes
    Bank = pd.read_excel(BankAct)
    Map = pd.read_excel(Mapping)
    Map["Closing Balance"] = None

    # Replace Nan with blank string
    Bank = Bank.fillna("")

    # Blank df for exceptions
    exdf = pd.DataFrame(None,
                        columns=["Bank Reference ID", "bank_closing_balance", "overnight_balance",
                                 "Calculated Closing Balance"])

    # Variable for keeping track of exceptions
    ex = False

    # Add columns to Bank df
    Bank2 = pd.DataFrame(None, columns=["Bank Reference ID", "Post Date", "Value Date", "Amount",
                                        "Description", "Bank Account", "Closing Balance", "Filename"])

    Bank = pd.concat([Bank, Bank2], axis=1)

    Bank["Bank Reference ID"] = Bank["Reference Number"]
    Bank["Post Date"] = pd.to_datetime(Bank["Cash Post Date"])
    Bank["Value Date"] = pd.to_datetime(Bank["Cash Value Date"])
    Bank["Amount"] = Bank["Transaction Amount Local"]

    for index, row in Bank.iterrows():
        trans = [row["Transaction Description 1"], row["Transaction Description 2"], row["Transaction Description 3"],
                 row["Transaction Description 4"], row["Transaction Description 5"], row["Transaction Description 6"],
                 row["Detailed Transaction Type Name"], row["Transaction Type"]]

        Bank.loc[index, "Description"] = " ".join(str(item) for item in trans)

    Bank["Bank Account"] = Bank["Cash Account Number"]
    Bank["Closing Balance"] = Bank["Closing Balance Local"]

    for index, row in Bank.iterrows():
        dt_string = datetime.now().strftime("%d-%m-%Y--%H-%M-%S")
        filename = str(row["Cash Account Number"]) + dt_string + ".csv"
        Bank.loc[index, "Filename"] = filename

    # Bank Ref ID df
    BRD = Map["Bank Ref ID"]

    # Bank Ref ID + Starting Balance df
    SB = Map[["Bank Ref ID", "Starting_Balance"]]

    for row in BRD:
        # Starting Balance
        bal = SB.loc[SB["Bank Ref ID"] == row]

        # Output dataframe
        out = Bank.loc[Bank["Cash Account Number"] == row]

        # MM Dataframe
        mm = out.loc[out["Description"].str.contains("STIF")]
        out = pd.concat([out, mm]).drop_duplicates(keep=False)

        # Create write_file and
        write_file = out[["Bank Reference ID", "Post Date", "Value Date", "Amount",
                          "Description", "Bank Account", "Closing Balance"]]

        write_file = write_file[write_file["Bank Reference ID"] != ""]

        if write_file.empty:
            print(str(row) + " has no activity")
            calc_closing_balance = 0

        else:
            bank_closing_balance = Bank[Bank["Bank Account"] == row]["Closing Balance"].iloc[-1]
            overnight_balance = mm["Amount"].sum()

            write_file["Bank Reference ID"] = "Starting Balance"
            write_file["Post Date"] = '2020-01-01'
            write_file["Value Date"] = '2020-01-01'
            write_file["Amount"] = bal["Starting_Balance"]
            write_file["Description"] = "Starting Balance"
            write_file["Bank Account"] = row
            write_file["Closing Balance"] = 0

            calc_closing_balance = write_file["Amount"].sum()

            if calc_closing_balance != bank_closing_balance and calc_closing_balance != overnight_balance:
                ex = True
                exlist = [row, bank_closing_balance, overnight_balance, calc_closing_balance]
                exdf.loc[len(exdf)] = exlist

            else:
                pass

            x_name = f"{row}-{dt_string}.xlsx"
            write_file.to_excel("./Output/" + x_name, sheet_name="Bank Transactions")
        Map.loc[Map["Bank Ref ID"] == row, 'Closing Balance'] = calc_closing_balance

    Map.to_excel("./Mapping/Cash_Rec_Mapping.xlsx")
    exdf.to_excel("./Output/exceptions.xlsx")


if __name__ == '__main__':
    Main()
