import decimal

import tinvest


class PortfolioUtil:

    def __init__(self, token: str) -> None:
        self.client = tinvest.SyncClient(token)
        self.portfolio = self.client.get_portfolio().payload.positions

    def get_currency_course(self, currency="USD") -> decimal.Decimal:
        supported_currencies = {'USD': dict(name='Доллар США', ticker=''),
                                'EURO': dict(name='Евро', ticker='')}
        if currency not in supported_currencies.keys():
            raise Exception(f"Unknown currency is given. "
                            f"Expected on of {supported_currencies}. Actual: {currency}")
        instruments = self.client.get_market_currencies().payload.instruments
        for i in instruments:
            if i.name == supported_currencies['USD']['name']:
                supported_currencies['USD']['ticker'] = i.ticker
            elif i.name == supported_currencies['EURO']['name']:
                supported_currencies['EURO']['ticker'] = i.ticker
            else:
                raise Exception(f"Unknown currency received from Tinkoff server. Received currency: {i.name}")
        currency_course = self.__get_paper_price(supported_currencies.get(currency)["ticker"], 1)
        return currency_course

    def get_portfolio_currencies(self) -> dict:
        portfolio_currencies = self.client.get_portfolio_currencies().payload.currencies
        sorted_currencies = dict()
        for i in portfolio_currencies:
            if i.balance == 0:
                continue
            sorted_currencies[i.currency.value] = i.balance
        return sorted_currencies

    def get_portfolio_currencies_in_rub(self) -> dict:
        currencies = self.get_portfolio_currencies()
        currencies_in_rub = {name: value*self.get_currency_course(name)
                             for name, value in currencies.items() if name != 'RUB'}
        if currencies.get('RUB'):
            currencies_in_rub['RUB'] = currencies.get('RUB')
        return currencies_in_rub

    def get_papers_prices(self) -> dict:
        paper_prices = dict(RUB=dict(), USD=dict(), EURO=dict())
        for name, lots_quantity, currency in self.__get_portfolio_papers():
            paper_price = self.__get_paper_price(name, lots_quantity)
            if paper_price == 0:
                continue
            paper_prices[currency][name] = paper_price
        return paper_prices

    def get_papers_prices_in_rub(self) -> dict:
        currencies = ['USD', 'EURO']
        paper_prices = self.get_papers_prices()
        for currency in currencies:
            tmp_papers = paper_prices[currency]
            paper_prices[currency] = dict()
            for name, price in tmp_papers.items():
                paper_prices['RUB'][name] = price * self.get_currency_course(currency)
        paper_prices = self.exclude_incorrect_papers(paper_prices)
        return paper_prices

    def __get_portfolio_papers(self) -> list:
        return [(i.ticker, i.balance, i.average_position_price.currency.value) for i in self.portfolio]

    def __get_paper_price(self, ticker, lots_quantity) -> decimal.Decimal:
        figi = self.client.get_market_search_by_ticker(ticker).payload.instruments[0].figi
        last_price = self.client.get_market_orderbook(figi, 1).payload.last_price
        if not last_price:
            raise Exception(f"Received incorrect paper price. Received: {last_price}")
        return last_price * lots_quantity

    def exclude_incorrect_papers(self, papers: dict) -> dict:
        exception_papers = ('USD000UTSTOM', )
        papers['RUB'] = {k: v for k, v in papers['RUB'].items() if k not in exception_papers}
        return papers
