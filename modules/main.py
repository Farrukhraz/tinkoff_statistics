import os
from modules.Utils import ExcelUtil, PortfolioUtil


def update_statistics() -> None:
    """ Update brokerage account statistics """
    portfolio = PortfolioUtil(os.environ.get('TOKEN'))
    print(portfolio.get_papers_prices())
    print(portfolio.get_currency_course("EURO"))
    print(portfolio.get_portfolio_currencies())

# print(portfolio.get_papers_prices_in_rub())
# print(portfolio.get_currency_course())
# print(portfolio.get_currency_course("USD"))
