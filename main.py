import gspread
import time
import typing
import requests
import json

RANDOM_CELL_LABEL: str = "A1"
LISTING_ID_RANGE: str = "A3:A1000"
STATUS_CELL_LABEL: str = "J1"


class CardListing(typing.NamedTuple):
    listing_id: int
    name: str
    color: str
    cmc: int
    type: str
    price: float
    url: str


def get_tcgplayer_data(listing_id) -> CardListing:
    """gets data from tcgplayer"""
    url: str = f"https://mp-search-api.tcgplayer.com/v1/product/{listing_id}/details?mpfev=1680"
    data: dict = requests.get(
        url,
    ).json() # gets data from url
    customAttrs: dict | None = data.get("customAttributes")  # check if key exists

    # creates CardListing object
    return CardListing(
        listing_id=listing_id,
        name=data.get("productName"),
        color=customAttrs.get("color") if customAttrs else None,
        cmc=customAttrs.get("convertedCost") if customAttrs else None,
        type=customAttrs.get("fullType") if customAttrs else None,
        price=data.get("marketPrice"),
        url=f"https://www.tcgplayer.com/product/{listing_id}",
    )


def connect() -> gspread.Worksheet:
    """Connects to google sheets and returns the first worksheet"""
    gc = gspread.service_account(filename="creds.json")
    spreadsheet: gspread.Spreadsheet = gc.open("MTG Card Collection")
    return spreadsheet.get_worksheet(0)


def write_row(sheet: gspread.Worksheet, cellRange: str, cardData: CardListing) -> None:
    """Writes the card data to a row"""
    cells: list[gspread.Cell] = sheet.range(cellRange)  # get the range as cells
    cardData = tuple(cardData)  # convert namedtuple to tuple
    for i, cell in enumerate(cells):
        # some filtering and assignment
        cell.value = (
            cardData[i] if not isinstance(cardData[i], list) else " ".join(cardData[i])
        )

    sheet.update_cells(cells)
    return None


def write_error(sheet: gspread.Worksheet, cell: gspread.Cell, message: str) -> None:
    """Writes an error message on the row that failed."""
    cell.value = f"ERROR: {message}"
    sheet.update_cells(
        [
            cell,
        ]
    )


def write_status(sheet: gspread.Worksheet, status: str):
    """Write the current status in a predefined cell"""
    sheet.update_acell(STATUS_CELL_LABEL, status)


def main():
    sheet: gspread.Worksheet = connect()

    randomCellOldValue: int = sheet.acell(RANDOM_CELL_LABEL).value

    while (
        True
    ):  # loop checks if the user has requested a reload, this is done by a cell changing value.
        try:
            newValueCheck: int = sheet.acell(RANDOM_CELL_LABEL).value
            if newValueCheck != randomCellOldValue:
                randomCellOldValue = newValueCheck
                write_status(sheet, status="Refreshing...")
                for listing in (
                    cell for cell in sheet.range(LISTING_ID_RANGE) if cell.value
                ):
                    try:
                        cdata = get_tcgplayer_data(listing.value)

                        write_row(
                            sheet=sheet,
                            cellRange=f"A{listing.row}:G{listing.row}",
                            cardData=cdata,
                        )
                    except Exception as e:
                        write_error(sheet, status=f"ERROR: {e}")

                    time.sleep(2)  # wait 2 seconds after each write
                write_status(sheet, status="Waiting...")
            else:
                write_status(sheet, status="Waiting...")
        except Exception as e:
            write_status(sheet, status=f"ERROR: {e}")
            raise e
        time.sleep(10)  # wait 10 seconds before checking if a refresh was requested


if __name__ == "__main__":
    main()
