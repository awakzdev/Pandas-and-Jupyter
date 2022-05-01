import pandas as pd


def inventory():
    """This will check whether the memory is available in the cabinet"""
    file = "S:\\Groups\\FV\\Memory\\everything_else\\MMEX\\InventoryFiles\\MMEX Inventory.xlsx"
    df = pd.ExcelFile(file).sheet_names

    # Filter sheets
    counter = 0
    sheets = []
    for sheet in df:
        if sheet == "EXTRA" or sheet == "Inventory Rules" or sheet == "Removed lines" or sheet == "EOL_Hynix_SODIMM" \
                or sheet == "EV" or sheet == "LPDDR4" or sheet == "LP4":
            pass
        else:
            counter += 1
            sheets.append(sheet)

    # Input will loop through sheets in IDC S/N column
    memory = input("Enter memory : ")

    # Compare and keep matching columns
    counter = 0
    for i in sheets:
        while i in sheets:
            if counter == len(sheets) - 1:
                break
            else:
                df = pd.read_excel(f"{file}", f"{sheets[counter]}")
                counter += 1
                a = df.columns
                b = ['ECC', 'Cabinet Qty', 'MMEX']
                keep_columns = [x for x in a if x in b]
                pd.set_option('display.max_columns', None)
                pd.set_option('display.width', None)
                pd.set_option('display.max_colwidth', None)
                df_out = df.loc[df.loc[:, 'IDC S/N'].fillna('nan').str.lower().str.contains(memory.lower()), :][
                    keep_columns]
                if df_out.empty:
                    pass
                else:
                    print(f"{sheets[counter]}\n" f"{df_out}")


while True:
    inventory()
