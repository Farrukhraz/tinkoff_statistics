import os
from modules.Utils import ExcelUtil, PortfolioUtil


def update_statistics() -> None:
    """ Update brokerage account statistics """
    portfolio = PortfolioUtil(os.environ.get('TOKEN'))
    excel_util = ExcelUtil('Invest.xlsx', portfolio)
    excel_util.update_papers_statistics()
