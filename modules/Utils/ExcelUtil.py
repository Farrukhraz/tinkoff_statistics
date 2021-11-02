import win32com.client

from typing import NamedTuple
from decimal import Decimal

from . import PortfolioUtil


class PapersTable:
    RUB, USD, EURO, *_ = (chr(i) for i in range(5, 10))
    WELL_KNOWN_CURRENCIES = dict(
        RUB=RUB,
        USD=USD,
        EURO=EURO,
    )
    paper_field = NamedTuple('PaperField', [('name', str), ('price_rub', Decimal), ('price_usd', Decimal)])
    balance_field = NamedTuple('BalanceField', [('name', str), ('currency_value', Decimal)])

    def __init__(self):
        self.table = {
            'title': 'Common',
            'body': [],
            'balance': []
        }
        self.__initialize_table()

    def append_paper(self, paper_name: str, paper_total_price: Decimal, currency: str = "RUB") -> None:
        """ Append paper to the table """
        pass

    def update_paper_value(self, paper_name: str, paper_total_price: Decimal) -> None:
        """ Update paper that exists in xlsx file. If paper is not in the table append it there """
        pass

    def delete_paper(self, paper_name: str) -> None:
        pass

    def update_balance(self, currency_name: str, currency_value: Decimal) -> None:
        if currency_name not in self.WELL_KNOWN_CURRENCIES:
            raise ValueError("Unknown currency is given")
        self.table['balance'].append(self.balance_field(currency_name, currency_value))

    def __initialize_table(self) -> None:
        """ Fill table with values from existing table in xlsx file """
        pass


class ExcelUtil:

    def __init__(self, file_path: str, portfolio: PortfolioUtil) -> None:
        self.portfolio: PortfolioUtil = portfolio
        self.file_path = file_path
        self.papers_range = range(1, 99)
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.papers_table = PapersTable()

    def update_papers_statistics(self) -> None:
        wb = self.excel.Workbooks.Open(self.file_path)
        sheet = wb.ActiveSheet
        try:
            paper_prices: dict = self.portfolio.get_papers_prices_in_rub().get('RUB')
            currencies = self.portfolio.get_portfolio_currencies_in_rub()
            paper_prices = {**paper_prices, **currencies}
        except AttributeError:
            raise AttributeError("Cannot update papers statistics 'cause cannot receive papers price")
        # for r in sheet.Range(f"A{self.papers_range[0]}:A{self.papers_range[1]}"):
        i = 0
        for i in self.papers_range:
            paper_name = sheet.Cells(i, 1).value
            if not paper_name:
                break
            paper_price = paper_prices.get(paper_name)
            if paper_price is not None:
                sheet.Cells(i, 2).value = round(paper_price)
                del paper_prices[paper_name]
        for name, price in paper_prices.items():
            sheet.Cells(i, 1).value = name
            sheet.Cells(i, 2).value = round(price)
            i += 1
        wb.Save()
        wb.Close()
        self.excel.Quit()


    def update_papers(self) -> None:
        currencies = self.portfolio.get_papers_prices()
        for currency, papers in currencies:
            for name, value in papers:
                self.papers_table.update_paper_value(name, value)

    def update_balance(self) -> None:
        currencies = self.portfolio.get_portfolio_currencies()
        rub = self.papers_table.WELL_KNOWN_CURRENCIES[self.papers_table.RUB]
        usd = self.papers_table.WELL_KNOWN_CURRENCIES[self.papers_table.USD]
        self.papers_table.update_balance(rub, currencies[rub])
        self.papers_table.update_balance(usd, currencies[usd])


