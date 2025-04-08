import pandas as pd
from datetime import date
import glob

price_indices = pd.read_excel("Prisindex.xlsx", decimal = ",")
input_files = glob.glob("In-filer/*.xlsx")

# Returns the `cost' at date `date' adjusted to today.
def adjust_cost(cost: float, date: date) -> float:
    global price_indices

    data_year = date.year

    yeardata = price_indices.loc[price_indices["Year"] >= data_year]["Percentage"]
    retval = cost * yeardata.prod()

    return round(retval, 2)


# Returns a list of BSAB.
def read_Xlsx_file(filename):

    df = pd.read_excel(filename, decimal = ",")
    sheet_length = len(df)

    # Skip to after "Kod"
    after_code = 0
    for i in range(sheet_length):
        row = df.iloc[i]
        if (row.iloc[0]) == "Kod":
            after_code = i + 1
            break

    def clean_str_cell(value):
        return value.strip() if isinstance(value, str) else ""


    BSABs = []

    project_name = df.iloc[0].iloc[1]
    project_date = date.fromisoformat(str(df.iloc[0].iloc[4])[0:10])
    code = ""
    description = ""

    for i in range(after_code, sheet_length): # For each row after Kod
        candidate_code = clean_str_cell(df.iloc[i].iloc[0]) # column "Kod"

        if candidate_code != "":
            code = candidate_code
            # For instance "Tillfällig gångbrygga"
            description = str(df.iloc[i].iloc[1])

        # code is now for instance "BCB.7122"

        unit = clean_str_cell(df.iloc[i].iloc[3]) # column "Enhet"

        price = df.iloc[i].iloc[5] # column "á-pris"

        # The files use different kinds of dashes.
        def is_dash(char: str) -> bool:
            stripped = ""

            # It might be a number.
            try:
                stripped = char.strip()
            except:
                return False

            return stripped == "-" or (len(stripped) == 1 and ord(stripped) == 8722)

        add = False
        amount = None

        if is_dash(unit) and is_dash(price):
            # This is a cost without quantity, some kind of event/moment

            # Column "Belopp"/"Kostnad"
            price = float('NaN')
            str_price = df.iloc[i].iloc[6]
            if not is_dash(str_price):
                price = float(str_price)

            # Column "Mängd"
            str_amount = df.iloc[i].iloc[4]
            if not is_dash(str_amount):
                amount = float(str_amount)

            add = True

        elif unit != "":
            amount = None

            try:
                amount = float(df.iloc[i].iloc[4]) # Column "Mängd"
            except:
                # We do nothing, amount is None
                pass

            add = True

        if add:
            BSABs.append((str(code),
                          str(description),
                          str(df.iloc[i].iloc[1]), # Column "Text"/sub_code
                          unit, # For instance "st" or "m".
                          amount,
                          price,
                          project_date,
                          adjust_cost(price, project_date),
                          project_name,
                          filename))
    return BSABs

# Returns a DataFrame of all the XSLX-files in `file_names' compiled into the
# format we use, BSAB.
def read_Xlsx_files(file_names):
    BSABs = []

    for f in file_names:
        o = read_Xlsx_file(f)
        BSABs += o

    #for d in BSABs:
        #print(d)

    BSABs.sort()

    df = pd.DataFrame(BSABs,
                      columns=['Kod',
                               'Beskrivning',
                               'Sub-kod',
                               'Enhet',
                               'Mängd',
                               'Á-pris',
                               'Datum',
                               'Justerat á-pris',
                               'Projekt',
                               'Filnamn'])

    return df

def to_excel(df, filename):
    # We pickle the DataFrames, but it seems Pandas doesn't support
    # pickling of style, so hence we make this specific to the XLSX export.
    df = df.style.map(lambda v: "background-color: #7DC4A6", subset=['Justerat á-pris'])

    df.to_excel(filename,
                index=False,
                freeze_panes=(1, 0))


def main() -> int:
    global input_files

    df = read_Xlsx_files(input_files)
    to_excel(df, "AMA-priser.xlsx")

    return 0


if __name__ == "__main__":
    main()
