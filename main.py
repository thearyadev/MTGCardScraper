import gspread
import time
import typing
import requests

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
    url: str = f"https://mpapi.tcgplayer.com/v2/product/{listing_id}/details"
    data: dict = requests.get(url).json()
    customAttrs: dict | None = data.get("customAttributes")

    return CardListing(
        listing_id=listing_id,
        name=data.get("productName"),
        color=customAttrs.get("color") if customAttrs else None,
        cmc=customAttrs.get("convertedCost") if customAttrs else None,
        type=customAttrs.get("fullType") if customAttrs else None,
        price=data.get("marketPrice"),
        url=f"https://www.tcgplayer.com/product/{listing_id}"
    )


def connect() -> gspread.Worksheet:
    gc = gspread.service_account(filename="creds.json")
    spreadsheet: gspread.Spreadsheet = gc.open("MTG Card Collection")
    return spreadsheet.get_worksheet(0)


def write_row(sheet: gspread.Worksheet, cellRange: str, cardData: CardListing) -> None:
    cells: list[gspread.Cell] = sheet.range(cellRange)
    cardData = tuple(cardData)
    for i, cell in enumerate(cells):
        cell.value = cardData[i] if not isinstance(cardData[i], list) else " ".join(cardData[i])

    sheet.update_cells(cells)
    return None


def write_error(sheet: gspread.Worksheet, cell: gspread.Cell, message: str) -> None:
    cell.value = f"ERROR: {message}"
    sheet.update_cells([cell, ])


def write_status(sheet: gspread.Worksheet, status: str):
    sheet.update_acell(STATUS_CELL_LABEL, status)


def main():
    sheet: gspread.Worksheet = connect()

    randomCellOldValue: int = sheet.acell(RANDOM_CELL_LABEL).value

    while True:
        try:
            newValueCheck: int = sheet.acell(RANDOM_CELL_LABEL).value
            if newValueCheck != randomCellOldValue:
                randomCellOldValue = newValueCheck
                write_status(sheet, status="Refreshing...")
                for listing in (cell for cell in sheet.range(LISTING_ID_RANGE) if cell.value):
                    try:
                        write_row(
                            sheet=sheet,
                            cellRange=f"A{listing.row}:G{listing.row}",
                            cardData=get_tcgplayer_data(listing.value)
                        )

                    except Exception as e:
                        write_error(sheet, sheet.cell(row=listing.row, col=2), message=str(e))
                    time.sleep(2)
                write_status(sheet, status="Waiting...")
            else:
                write_status(sheet, status="Waiting...")
        except Exception as e:
            write_status(sheet, status=f"ERROR: {e}")
        time.sleep(10)


if __name__ == '__main__':
    main()
