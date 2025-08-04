"""Creation of controller list"""

from . import balance_analysis

routes = [
    balance_analysis.router,
]

tags = [
    *balance_analysis.tags,
]
